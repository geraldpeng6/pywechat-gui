from __future__ import annotations

import re
from pathlib import Path
from typing import Callable

from .adapter import PyWeixinAdapter
from .import_export import dump_table
from .models import ChatBatchExportRequest, ChatBatchExportResult, ChatExportRequest, ChatExportResult, RuntimeOptions
from .paths import create_unique_timestamped_dir


ProgressCallback = Callable[[str], None]
StopCallback = Callable[[], bool]


def _sanitize_filename(value: str) -> str:
    cleaned = re.sub(r'[\\/:*?"<>|]+', "_", value).strip()
    return cleaned or "会话导出"


MEDIA_FOLDER_NAME = "聊天图片与视频"


class ChatExportService:
    def __init__(self, adapter: PyWeixinAdapter):
        self.adapter = adapter

    def export_chat_bundle(
        self,
        request: ChatExportRequest,
        runtime_options: RuntimeOptions,
        on_progress: ProgressCallback | None = None,
        should_stop: StopCallback | None = None,
    ) -> ChatExportResult:
        self._check_stop(should_stop)
        target_root = Path(request.target_folder).expanduser().resolve()
        export_folder = create_unique_timestamped_dir(target_root, _sanitize_filename(request.session_name))

        result = ChatExportResult(session_name=request.session_name, export_folder=str(export_folder))
        warnings: list[str] = []

        if request.export_messages:
            if on_progress:
                on_progress("正在导出聊天消息...")
            messages, timestamps = self.adapter.dump_chat_history(
                session_name=request.session_name,
                number=request.message_limit,
                options=runtime_options,
            )
            ordered_messages = self._ordered_message_rows(messages, timestamps)
            result.message_count = len(ordered_messages)
            xlsx_path = export_folder / "聊天记录.xlsx"
            dump_table(
                ["序号", "时间", "消息内容"],
                [
                    {
                        "序号": row["index"],
                        "时间": row["timestamp"],
                        "消息内容": row["message"],
                    }
                    for row in ordered_messages
                ],
                xlsx_path,
            )
            result.messages_xlsx = str(xlsx_path)

        self._check_stop(should_stop)

        if request.export_files:
            if on_progress:
                on_progress("正在导出聊天文件...")
            files_folder = export_folder / "聊天文件"
            files_folder.mkdir(parents=True, exist_ok=True)
            result.files_folder = str(files_folder)
            try:
                self.adapter.save_chat_files(
                    session_name=request.session_name,
                    number=request.file_limit,
                    target_folder=str(files_folder),
                    options=runtime_options,
                )
                exported_files = [path for path in files_folder.iterdir() if path.is_file()]
                result.file_count = len(exported_files)
            except Exception as exc:
                if not self._looks_like_chat_file_parse_error(exc):
                    raise
                warnings.append(
                    "聊天文件未整理到结果中。"
                    "该会话可能暂时没有可保存的聊天文件，或当前页面未能成功读取到文件列表。"
                )

        self._check_stop(should_stop)

        if request.export_images:
            if on_progress:
                on_progress("正在导出聊天图片/视频...")
            media_folder = export_folder / MEDIA_FOLDER_NAME
            media_folder.mkdir(parents=True, exist_ok=True)
            result.media_folder = str(media_folder)
            exported_media = self.adapter.save_chat_media(
                session_name=request.session_name,
                number=request.file_limit,
                target_folder=str(media_folder),
                options=runtime_options,
            )
            result.media_count = len([path for path in exported_media if Path(path).is_file()])

        result.warnings = warnings
        return result

    def export_chat_batch(
        self,
        request: ChatBatchExportRequest,
        runtime_options: RuntimeOptions,
        on_progress: ProgressCallback | None = None,
        should_stop: StopCallback | None = None,
    ) -> ChatBatchExportResult:
        target_root = Path(request.target_folder).expanduser().resolve()
        batch_root = create_unique_timestamped_dir(target_root, "批量会话导出")

        session_results: list[ChatExportResult] = []
        failed_sessions: list[dict[str, str]] = []
        session_names = [name.strip() for name in request.session_names if name.strip()]

        for session_name in session_names:
            self._check_stop(should_stop)
            if on_progress:
                on_progress(f"正在导出会话：{session_name}")
            try:
                single_request = ChatExportRequest(
                    session_name=session_name,
                    target_folder=str(batch_root),
                    export_messages=request.export_messages,
                    export_files=request.export_files,
                    export_images=request.export_images,
                    message_limit=request.message_limit,
                    file_limit=request.file_limit,
                )
                result = self.export_chat_bundle(
                    request=single_request,
                    runtime_options=runtime_options,
                    on_progress=None,
                    should_stop=should_stop,
                )
                session_results.append(result)
            except Exception as exc:
                failed_sessions.append({"session_name": session_name, "error": str(exc)})

        summary = ChatBatchExportResult(
            export_root=str(batch_root),
            total_sessions=len(session_names),
            success_count=len(session_results),
            failure_count=len(failed_sessions),
            session_results=session_results,
            failed_sessions=failed_sessions,
        )
        return summary

    @staticmethod
    def _ordered_message_rows(messages: list[str], timestamps: list[str]) -> list[dict[str, str | int]]:
        rows = [
            {
                "index": index + 1,
                "timestamp": timestamps[index] if index < len(timestamps) else "",
                "message": message,
            }
            for index, message in enumerate(messages)
        ]
        rows.reverse()
        for index, row in enumerate(rows, start=1):
            row["index"] = index
        return rows

    @staticmethod
    def _build_summary_text(result: ChatExportResult) -> str:
        lines = [
            f"会话名称：{result.session_name}",
            f"导出目录：{result.export_folder}",
            f"消息数量：{result.message_count}",
            f"文件数量：{result.file_count}",
            f"图片/视频数量：{result.media_count}",
        ]
        if result.messages_xlsx:
            lines.append(f"聊天记录：{result.messages_xlsx}")
        if result.files_folder:
            lines.append(f"文件目录：{result.files_folder}")
        if result.media_folder:
            lines.append(f"图片/视频目录：{result.media_folder}")
        if result.warnings:
            lines.append("注意事项：")
            lines.extend(f"- {warning}" for warning in result.warnings)
        return "\n".join(lines)

    @staticmethod
    def _check_stop(should_stop: StopCallback | None) -> None:
        if should_stop and should_stop():
            raise RuntimeError("导出任务已停止")

    @staticmethod
    def _looks_like_chat_file_parse_error(exc: Exception) -> bool:
        message = str(exc)
        return isinstance(exc, AttributeError) and "'NoneType' object has no attribute 'group'" in message
