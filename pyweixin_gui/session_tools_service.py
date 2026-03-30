from __future__ import annotations

from typing import Callable

from .adapter import PyWeixinAdapter
from .models import GroupScanResult, RuntimeOptions, SessionScanRequest, SessionScanResult


ProgressCallback = Callable[[str], None]


class SessionToolsService:
    def __init__(self, adapter: PyWeixinAdapter):
        self.adapter = adapter

    def scan_sessions(
        self,
        request: SessionScanRequest,
        runtime_options: RuntimeOptions,
        on_progress: ProgressCallback | None = None,
    ) -> SessionScanResult:
        if on_progress:
            on_progress("正在采集微信会话列表...")
        rows = self.adapter.dump_sessions(
            chatted_only=request.chatted_only,
            options=runtime_options,
        )
        return SessionScanResult(rows=rows)

    def scan_groups(
        self,
        runtime_options: RuntimeOptions,
        on_progress: ProgressCallback | None = None,
    ) -> GroupScanResult:
        if on_progress:
            on_progress("正在采集群聊列表...")
        rows = self.adapter.dump_groups(runtime_options)
        return GroupScanResult(rows=rows)
