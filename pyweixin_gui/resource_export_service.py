from __future__ import annotations

from datetime import datetime
from pathlib import Path
from typing import Callable

from .adapter import PyWeixinAdapter
from .models import ResourceExportKind, ResourceExportRequest, ResourceExportResult, RuntimeOptions


ProgressCallback = Callable[[str], None]
StopCallback = Callable[[], bool]


def export_kind_label(kind: ResourceExportKind) -> str:
    mapping = {
        ResourceExportKind.RECENT_FILES: "最近聊天文件",
        ResourceExportKind.WXFILES: "微信聊天文件",
        ResourceExportKind.VIDEOS: "微信聊天视频",
    }
    return mapping[kind]


class ResourceExportService:
    def __init__(self, adapter: PyWeixinAdapter):
        self.adapter = adapter

    def run_export(
        self,
        request: ResourceExportRequest,
        runtime_options: RuntimeOptions,
        on_progress: ProgressCallback | None = None,
        should_stop: StopCallback | None = None,
    ) -> ResourceExportResult:
        self._check_stop(should_stop)
        target_folder = Path(request.target_folder).expanduser().resolve()
        target_folder.mkdir(parents=True, exist_ok=True)

        if request.export_kind is ResourceExportKind.RECENT_FILES:
            if on_progress:
                on_progress("正在导出最近聊天文件...")
            exported = self.adapter.export_recent_files(str(target_folder), runtime_options)
        elif request.export_kind is ResourceExportKind.WXFILES:
            if on_progress:
                on_progress("正在导出微信聊天文件...")
            exported = self.adapter.export_wxfiles(request.year, request.month or None, str(target_folder))
        else:
            if on_progress:
                on_progress("正在导出微信聊天视频...")
            exported = self.adapter.export_videos(request.year, request.month or None, str(target_folder))

        self._check_stop(should_stop)
        exported_paths = [str(path) for path in exported]
        result = ResourceExportResult(
            export_kind=request.export_kind,
            target_folder=str(target_folder),
            exported_count=len(exported_paths),
            exported_paths=exported_paths,
        )
        summary_path = target_folder / f"{request.export_kind.value}-summary-{datetime.now().strftime('%Y%m%d-%H%M%S')}.txt"
        summary_path.write_text(self._build_summary(result, request), encoding="utf-8")
        result.summary_txt = str(summary_path)
        return result

    @staticmethod
    def _build_summary(result: ResourceExportResult, request: ResourceExportRequest) -> str:
        lines = [
            f"导出类型：{export_kind_label(result.export_kind)}",
            f"导出目录：{result.target_folder}",
            f"导出数量：{result.exported_count}",
            f"年份：{request.year}",
            f"月份：{request.month or '全部'}",
        ]
        if result.exported_paths:
            lines.append("导出结果预览：")
            lines.extend(f"- {path}" for path in result.exported_paths[:10])
            if len(result.exported_paths) > 10:
                lines.append(f"- ... 共 {len(result.exported_paths)} 项")
        return "\n".join(lines)

    @staticmethod
    def _check_stop(should_stop: StopCallback | None) -> None:
        if should_stop and should_stop():
            raise RuntimeError("导出任务已停止")
