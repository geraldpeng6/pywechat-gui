from __future__ import annotations

import csv
import json
import re
from dataclasses import asdict
from datetime import datetime
from pathlib import Path
from typing import Callable

from .adapter import PyWeixinAdapter
from .models import ChatExportRequest, ChatExportResult, RuntimeOptions


ProgressCallback = Callable[[str], None]
StopCallback = Callable[[], bool]


def _sanitize_filename(value: str) -> str:
    cleaned = re.sub(r'[\\/:*?"<>|]+', "_", value).strip()
    return cleaned or "chat-export"


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
        timestamp = datetime.now().strftime("%Y%m%d-%H%M%S")
        export_folder = target_root / f"{_sanitize_filename(request.session_name)}-{timestamp}"
        export_folder.mkdir(parents=True, exist_ok=True)

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
            result.message_count = len(messages)
            csv_path = export_folder / "messages.csv"
            json_path = export_folder / "messages.json"
            self._write_messages_csv(csv_path, messages, timestamps)
            json_path.write_text(
                json.dumps(
                    [
                        {"index": index + 1, "timestamp": timestamps[index] if index < len(timestamps) else "", "message": message}
                        for index, message in enumerate(messages)
                    ],
                    ensure_ascii=False,
                    indent=2,
                ),
                encoding="utf-8",
            )
            result.messages_csv = str(csv_path)
            result.messages_json = str(json_path)

        self._check_stop(should_stop)

        if request.export_files:
            if on_progress:
                on_progress("正在导出聊天文件...")
            files_folder = export_folder / "files"
            files_folder.mkdir(parents=True, exist_ok=True)
            self.adapter.save_chat_files(
                session_name=request.session_name,
                number=request.file_limit,
                target_folder=str(files_folder),
                options=runtime_options,
            )
            exported_files = [path for path in files_folder.iterdir() if path.is_file()]
            result.file_count = len(exported_files)
            result.files_folder = str(files_folder)

        if request.export_images:
            warnings.append("当前库未提供稳定的历史图片/视频一键导出能力，本次已跳过图片与视频。")

        result.warnings = warnings
        summary_json = export_folder / "export-summary.json"
        summary_txt = export_folder / "export-summary.txt"
        summary_json.write_text(json.dumps(asdict(result), ensure_ascii=False, indent=2), encoding="utf-8")
        summary_txt.write_text(self._build_summary_text(result), encoding="utf-8")
        result.summary_json = str(summary_json)
        result.summary_txt = str(summary_txt)
        return result

    @staticmethod
    def _write_messages_csv(path: Path, messages: list[str], timestamps: list[str]) -> None:
        with path.open("w", encoding="utf-8-sig", newline="") as handle:
            writer = csv.DictWriter(handle, fieldnames=["index", "timestamp", "message"])
            writer.writeheader()
            for index, message in enumerate(messages):
                writer.writerow(
                    {
                        "index": index + 1,
                        "timestamp": timestamps[index] if index < len(timestamps) else "",
                        "message": message,
                    }
                )

    @staticmethod
    def _build_summary_text(result: ChatExportResult) -> str:
        lines = [
            f"会话名称：{result.session_name}",
            f"导出目录：{result.export_folder}",
            f"消息数量：{result.message_count}",
            f"文件数量：{result.file_count}",
        ]
        if result.messages_csv:
            lines.append(f"消息 CSV：{result.messages_csv}")
        if result.messages_json:
            lines.append(f"消息 JSON：{result.messages_json}")
        if result.files_folder:
            lines.append(f"文件目录：{result.files_folder}")
        if result.warnings:
            lines.append("注意事项：")
            lines.extend(f"- {warning}" for warning in result.warnings)
        return "\n".join(lines)

    @staticmethod
    def _check_stop(should_stop: StopCallback | None) -> None:
        if should_stop and should_stop():
            raise RuntimeError("导出任务已停止")
