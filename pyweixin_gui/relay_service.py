from __future__ import annotations

from datetime import datetime
import json
from pathlib import Path
import re
from typing import Callable

from .adapter import PyWeixinAdapter
from .models import (
    RelayCollectFilesRequest,
    RelayCollectionResult,
    RelayCollectTextRequest,
    RelayItemType,
    RelayPackageRow,
    RelayRouteRow,
    RelaySendRequest,
    RelaySendResult,
    RelayTargetResult,
    RelayValidationRequest,
    RelayValidationResult,
    RuntimeOptions,
)
from .paths import ensure_app_dirs


ProgressCallback = Callable[[str], None]
StopCallback = Callable[[], bool]

IMAGE_SUFFIXES = {".png", ".jpg", ".jpeg", ".gif", ".bmp", ".webp", ".tiff"}


class RelayService:
    def __init__(self, adapter: PyWeixinAdapter):
        self.adapter = adapter

    def collect_text_rows(
        self,
        request: RelayCollectTextRequest,
        runtime_options: RuntimeOptions,
        on_progress: ProgressCallback | None = None,
    ) -> RelayCollectionResult:
        errors = request.validate()
        if errors:
            raise ValueError(next(iter(errors.values())))
        if on_progress:
            on_progress("正在采集上游文本消息...")
        messages, timestamps = self.adapter.dump_chat_history(
            session_name=request.source_session,
            number=request.message_limit,
            options=runtime_options,
        )
        now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        rows = [
            RelayPackageRow(
                sequence=index + 1,
                item_type=RelayItemType.TEXT,
                source_session=request.source_session,
                content=message,
                collected_at=timestamps[index] if index < len(timestamps) and timestamps[index] else now,
            )
            for index, message in enumerate(messages)
            if str(message or "").strip()
        ]
        warning = "" if rows else "当前没有采集到可转发的文本消息。"
        return RelayCollectionResult(source_session=request.source_session, rows=rows, warning=warning)

    def collect_file_rows(
        self,
        request: RelayCollectFilesRequest,
        runtime_options: RuntimeOptions,
        on_progress: ProgressCallback | None = None,
    ) -> RelayCollectionResult:
        errors = request.validate()
        if errors:
            raise ValueError(next(iter(errors.values())))
        if on_progress:
            on_progress("正在采集上游聊天文件...")
        cache_dir = self._relay_cache_dir(request.source_session)
        cache_dir.mkdir(parents=True, exist_ok=True)
        warning = ""
        try:
            self.adapter.save_chat_files(
                session_name=request.source_session,
                number=request.file_limit,
                target_folder=str(cache_dir),
                options=runtime_options,
            )
        except Exception as exc:
            if not self._looks_like_chat_file_parse_error(exc):
                raise
            warning = (
                "当前未能从微信聊天文件列表识别出结果数量。"
                "这通常表示该会话暂无可导出的聊天文件，或当前微信界面文案与 pyweixin 的解析逻辑不匹配。"
            )
        files = sorted([path for path in cache_dir.iterdir() if path.is_file()], key=lambda item: item.stat().st_mtime, reverse=True)
        rows = []
        for index, file_path in enumerate(files[: request.file_limit], start=1):
            item_type = RelayItemType.IMAGE if file_path.suffix.lower() in IMAGE_SUFFIXES else RelayItemType.FILE
            rows.append(
                RelayPackageRow(
                    sequence=index,
                    item_type=item_type,
                    source_session=request.source_session,
                    content=file_path.name,
                    file_path=str(file_path),
                    collected_at=datetime.fromtimestamp(file_path.stat().st_mtime).strftime("%Y-%m-%d %H:%M:%S"),
                )
            )
        if not rows and not warning:
            warning = "当前没有采集到可转发的聊天文件。"
        return RelayCollectionResult(source_session=request.source_session, rows=rows, warning=warning)

    def validate_routes(
        self,
        request: RelayValidationRequest,
        runtime_options: RuntimeOptions,
        on_progress: ProgressCallback | None = None,
    ) -> RelayValidationResult:
        errors = request.validate()
        if errors:
            raise ValueError(next(iter(errors.values())))
        checked_count = 0
        found_count = 0
        missing_count = 0
        updated_rows: list[RelayRouteRow] = []
        matched_rows = self._matched_route_rows(request.route_rows, request.source_session)
        for row in request.route_rows:
            new_row = RelayRouteRow.from_mapping(row.__dict__)
            if row not in matched_rows or not row.enabled:
                updated_rows.append(new_row)
                continue
            checked_count += 1
            if on_progress:
                on_progress(f"正在验证下游会话：{row.downstream_session}")
            try:
                self.adapter.validate_session(row.downstream_session, runtime_options)
                new_row.validation_status = "已找到"
                new_row.validation_message = "会话可打开"
                found_count += 1
            except Exception as exc:
                ui_error = self.adapter.map_runtime_exception(exc)
                new_row.validation_status = "未找到"
                new_row.validation_message = ui_error.message or ui_error.title
                missing_count += 1
            updated_rows.append(new_row)
        return RelayValidationResult(
            source_session=request.source_session,
            route_rows=updated_rows,
            checked_count=checked_count,
            found_count=found_count,
            missing_count=missing_count,
        )

    def load_export_folder_rows(self, folder_path: str | Path) -> RelayCollectionResult:
        folder = Path(folder_path).expanduser().resolve()
        if not folder.is_dir():
            raise ValueError("所选导出目录不存在。")
        source_session = folder.name
        summary_json = folder / "export-summary.json"
        if summary_json.exists():
            try:
                summary = json.loads(summary_json.read_text(encoding="utf-8"))
                source_session = str(summary.get("session_name", source_session)).strip() or source_session
            except json.JSONDecodeError:
                pass

        rows: list[RelayPackageRow] = []
        messages_json = folder / "messages.json"
        if messages_json.exists():
            payload = json.loads(messages_json.read_text(encoding="utf-8"))
            if isinstance(payload, list):
                for index, item in enumerate(payload, start=1):
                    if not isinstance(item, dict):
                        continue
                    message = str(item.get("message", "") or "").strip()
                    if not message:
                        continue
                    rows.append(
                        RelayPackageRow(
                            sequence=len(rows) + 1,
                            item_type=RelayItemType.TEXT,
                            source_session=source_session,
                            content=message,
                            collected_at=str(item.get("timestamp", "") or "").strip(),
                        )
                    )
        files_folder = folder / "files"
        if files_folder.exists():
            file_items = sorted([path for path in files_folder.iterdir() if path.is_file()], key=lambda item: item.stat().st_mtime, reverse=True)
            for path in file_items:
                item_type = RelayItemType.IMAGE if path.suffix.lower() in IMAGE_SUFFIXES else RelayItemType.FILE
                rows.append(
                    RelayPackageRow(
                        sequence=len(rows) + 1,
                        item_type=item_type,
                        source_session=source_session,
                        content=path.name,
                        file_path=str(path),
                        collected_at=datetime.fromtimestamp(path.stat().st_mtime).strftime("%Y-%m-%d %H:%M:%S"),
                    )
                )
        warning = "" if rows else "导出目录里没有识别到可转发的文本或文件。"
        return RelayCollectionResult(source_session=source_session, rows=rows, warning=warning)

    def send_package(
        self,
        request: RelaySendRequest,
        runtime_options: RuntimeOptions,
        on_progress: ProgressCallback | None = None,
        should_stop: StopCallback | None = None,
    ) -> RelaySendResult:
        errors = request.validate()
        if errors:
            raise ValueError(next(iter(errors.values())))
        package_rows = [item for item in sorted(request.package_rows, key=lambda item: item.sequence) if item.enabled]
        if not package_rows:
            raise ValueError("请至少勾选一条转发内容。")
        if request.test_only:
            targets = ["文件传输助手"]
        else:
            matched = self._matched_route_rows(request.route_rows, request.source_session)
            targets = self._unique_targets([row.downstream_session for row in matched if row.enabled and row.downstream_session.strip()])
            if not targets:
                raise ValueError("当前上游没有可用的下游路由。")
        results: list[RelayTargetResult] = []
        success_count = 0
        failure_count = 0
        for target_index, target_session in enumerate(targets, start=1):
            if should_stop and should_stop():
                break
            sent_count = 0
            try:
                for item in package_rows:
                    if on_progress:
                        on_progress(f"正在发送 {target_index}/{len(targets)}：{target_session} -> {item.preview_text()[:20]}")
                    self.adapter.send_relay_item(target_session, item, runtime_options)
                    sent_count += 1
                results.append(RelayTargetResult(target_session=target_session, success=True, sent_count=sent_count))
                success_count += 1
            except Exception as exc:
                ui_error = self.adapter.map_runtime_exception(exc)
                results.append(
                    RelayTargetResult(
                        target_session=target_session,
                        success=False,
                        sent_count=sent_count,
                        error_message=ui_error.message or ui_error.title,
                    )
                )
                failure_count += 1
        return RelaySendResult(
            source_session=request.source_session,
            test_only=request.test_only,
            item_count=len(package_rows),
            target_count=len(targets),
            success_count=success_count,
            failure_count=failure_count,
            results=results,
        )

    @staticmethod
    def _matched_route_rows(route_rows: list[RelayRouteRow], source_session: str) -> list[RelayRouteRow]:
        return [row for row in route_rows if row.upstream_session.strip() == source_session.strip()]

    @staticmethod
    def keep_latest_file_rows(rows: list[RelayPackageRow]) -> list[RelayPackageRow]:
        latest_by_key: dict[str, tuple[float, int]] = {}
        for index, row in enumerate(rows):
            if row.item_type not in {RelayItemType.FILE, RelayItemType.IMAGE} or not row.file_path:
                continue
            key = RelayService._normalized_file_key(row.file_path)
            score = RelayService._file_score(row)
            existing = latest_by_key.get(key)
            if existing is None or score > existing[0]:
                latest_by_key[key] = (score, index)
        updated: list[RelayPackageRow] = []
        for index, row in enumerate(rows):
            copied = RelayPackageRow.from_mapping(row.__dict__ | {"item_type": row.item_type.value})
            if row.item_type in {RelayItemType.FILE, RelayItemType.IMAGE} and row.file_path:
                key = RelayService._normalized_file_key(row.file_path)
                latest_index = latest_by_key.get(key, (-1, -1))[1]
                if latest_index != index:
                    copied.enabled = False
                    copied.remark = "旧版本，已自动取消勾选"
            updated.append(copied)
        return updated

    @staticmethod
    def _unique_targets(targets: list[str]) -> list[str]:
        seen: set[str] = set()
        result: list[str] = []
        for target in targets:
            if target not in seen:
                seen.add(target)
                result.append(target)
        return result

    @staticmethod
    def _looks_like_chat_file_parse_error(exc: Exception) -> bool:
        return isinstance(exc, AttributeError) and "'NoneType' object has no attribute 'group'" in str(exc)

    @staticmethod
    def _relay_cache_dir(source_session: str) -> Path:
        dirs = ensure_app_dirs()
        safe_name = "".join(char if char.isalnum() or char in {"-", "_"} else "_" for char in source_session).strip("_") or "relay-source"
        timestamp = datetime.now().strftime("%Y%m%d-%H%M%S")
        return dirs["local_dir"] / "relay-cache" / f"{safe_name}-{timestamp}"

    @staticmethod
    def _normalized_file_key(file_path: str) -> str:
        path = Path(file_path)
        stem = re.sub(r"\(\d+\)$", "", path.stem)
        return f"{stem.lower()}{path.suffix.lower()}"

    @staticmethod
    def _file_score(row: RelayPackageRow) -> float:
        try:
            return Path(row.file_path).stat().st_mtime
        except OSError:
            return float(row.sequence)
