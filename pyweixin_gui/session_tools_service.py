from __future__ import annotations

from typing import Callable

from .adapter import PyWeixinAdapter
from .models import GroupMembersRequest, GroupMembersResult, GroupScanResult, RuntimeOptions, SessionScanRequest, SessionScanResult


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
            no_official=request.no_official,
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

    def load_group_members(
        self,
        request: GroupMembersRequest,
        runtime_options: RuntimeOptions,
        on_progress: ProgressCallback | None = None,
    ) -> GroupMembersResult:
        errors = request.validate()
        if errors:
            raise ValueError(next(iter(errors.values())))
        if on_progress:
            on_progress(f"正在读取群成员：{request.group_name}")
        rows = self.adapter.dump_group_members(request.group_name, runtime_options)
        return GroupMembersResult(group_name=request.group_name, member_count=len(rows), rows=rows)
