from __future__ import annotations

from collections import Counter
from datetime import datetime
import json
from pathlib import Path
import re
import shutil
from typing import Callable

from .adapter import PyWeixinAdapter
from .import_export import dump_table, load_table_rows
from .models import (
    RelayCollectFilesRequest,
    RelayCollectMediaRequest,
    RelayCollectionResult,
    RelayCollectTextRequest,
    RelayItemType,
    RelayPackageExportRequest,
    RelayPackageExportResult,
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
MEDIA_PLACEHOLDER_LABELS = {
    "图片": "图片",
    "[图片]": "图片",
    "视频": "视频",
    "[视频]": "视频",
    "动画表情": "动画表情",
    "[动画表情]": "动画表情",
}


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
            on_progress("正在采集来源会话里的文字消息...")
        messages, timestamps = self.adapter.dump_chat_history(
            session_name=request.source_session,
            number=request.message_limit,
            options=runtime_options,
        )
        now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        skipped_media: Counter[str] = Counter()
        text_entries = []
        for index, message in enumerate(messages):
            message_text = str(message or "").strip()
            if not message_text:
                continue
            placeholder_label = self._placeholder_media_label(message_text)
            if placeholder_label:
                skipped_media[placeholder_label] += 1
                continue
            text_entries.append(
                (
                    message_text,
                    timestamps[index] if index < len(timestamps) and timestamps[index] else now,
                )
            )
        # pyweixin 返回的聊天记录顺序是“从晚到早”，这里反转成普通聊天阅读顺序。
        text_entries.reverse()
        rows = [
            RelayPackageRow(
                sequence=index,
                item_type=RelayItemType.TEXT,
                source_session=request.source_session,
                content=message,
                collected_at=timestamp,
            )
            for index, (message, timestamp) in enumerate(text_entries, start=1)
        ]
        warning = "" if rows else "当前没有采集到可转发的文本消息。"
        if skipped_media:
            media_summary = "、".join(f"{label}{count}条" for label, count in skipped_media.items())
            extra = f"已自动跳过 {media_summary}。如需把这些内容继续加入发送列表，可再点“采集图片/视频”或手动补入本地素材。"
            warning = f"{warning} {extra}".strip()
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
                "这次没有整理到聊天文件。"
                "该会话可能暂时没有可保存的文件，或当前页面未能成功读取到文件列表。"
            )
        files = sorted([path for path in cache_dir.iterdir() if path.is_file()], key=lambda item: item.stat().st_mtime)
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

    def collect_media_rows(
        self,
        request: RelayCollectMediaRequest,
        runtime_options: RuntimeOptions,
        on_progress: ProgressCallback | None = None,
    ) -> RelayCollectionResult:
        errors = request.validate()
        if errors:
            raise ValueError(next(iter(errors.values())))
        if on_progress:
            on_progress("正在采集聊天记录中的图片/视频...")
        cache_dir = self._relay_cache_dir(request.source_session)
        cache_dir.mkdir(parents=True, exist_ok=True)
        saved_paths = self.adapter.save_chat_media(
            session_name=request.source_session,
            number=request.media_limit,
            target_folder=str(cache_dir),
            options=runtime_options,
        )
        files = [Path(path) for path in reversed(saved_paths) if Path(path).is_file()]
        rows: list[RelayPackageRow] = []
        for index, file_path in enumerate(files, start=1):
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
        warning = "" if rows else "当前没有采集到可转发的历史图片或视频。"
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
        for row in request.route_rows:
            new_row = RelayRouteRow.from_mapping(row.__dict__)
            if not row.downstream_session.strip():
                updated_rows.append(new_row)
                continue
            checked_count += 1
            if on_progress:
                on_progress(f"正在验证收件人：{row.downstream_session}")
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

    def load_folder_rows(self, folder_path: str | Path) -> RelayCollectionResult:
        folder = Path(folder_path).expanduser().resolve()
        if not folder.is_dir():
            raise ValueError("所选文件夹不存在。")
        relay_manifest_xlsx = folder / "发送清单.xlsx"
        if not relay_manifest_xlsx.exists():
            relay_manifest_xlsx = folder / "转发清单.xlsx"
        if relay_manifest_xlsx.exists():
            return self._load_relay_package_rows(folder, relay_manifest_xlsx)
        relay_manifest = folder / "relay-package.json"
        if relay_manifest.exists():
            return self._load_relay_package_rows(folder, relay_manifest)
        return self._load_export_folder_rows(folder)

    def load_export_folder_rows(self, folder_path: str | Path) -> RelayCollectionResult:
        return self.load_folder_rows(folder_path)

    def export_package_folder(self, request: RelayPackageExportRequest) -> RelayPackageExportResult:
        errors = request.validate()
        if errors:
            raise ValueError(next(iter(errors.values())))
        target_root = Path(request.target_folder).expanduser().resolve()
        target_root.mkdir(parents=True, exist_ok=True)
        package_name = request.package_name.strip() or request.source_session.strip() or "发送任务"
        timestamp = datetime.now().strftime("%Y%m%d-%H%M%S")
        package_folder = target_root / f"{self._sanitize_name(package_name)}-{timestamp}"
        package_folder.mkdir(parents=True, exist_ok=True)
        files_folder = package_folder / "素材文件"
        files_folder.mkdir(parents=True, exist_ok=True)

        ordered_rows = list(sorted(request.package_rows, key=lambda item: item.sequence))
        manifest_items: list[dict[str, str | int]] = []
        message_count = 0
        file_count = 0

        for index, row in enumerate(ordered_rows, start=1):
            item_mapping: dict[str, str | int] = {
                "序号": index,
                "来源会话": request.source_session,
                "类型": self._item_type_label(row.item_type),
                "内容预览": row.content,
                "采集时间": row.collected_at,
                "相对路径": "",
            }
            if row.item_type is RelayItemType.TEXT:
                message_count += 1
            else:
                source_path = Path(row.file_path).expanduser().resolve()
                copied_path = self._copy_asset_to_folder(source_path, files_folder, index)
                item_mapping["相对路径"] = str(Path("素材文件") / copied_path.name)
                item_mapping["内容预览"] = copied_path.name
                file_count += 1
            manifest_items.append(item_mapping)

        manifest_path = package_folder / "发送清单.xlsx"
        dump_table(["序号", "来源会话", "类型", "内容预览", "采集时间", "相对路径"], manifest_items, manifest_path)
        return RelayPackageExportResult(
            source_session=request.source_session,
            package_name=package_name,
            package_folder=str(package_folder),
            item_count=len(ordered_rows),
            message_count=message_count,
            file_count=file_count,
            manifest_path=str(manifest_path),
            files_folder=str(files_folder),
        )

    def _load_export_folder_rows(self, folder: Path) -> RelayCollectionResult:
        source_session = self._guess_export_session_name(folder)
        summary_json = folder / "export-summary.json"
        if summary_json.exists():
            try:
                summary = json.loads(summary_json.read_text(encoding="utf-8"))
                source_session = str(summary.get("session_name", source_session)).strip() or source_session
            except json.JSONDecodeError:
                pass

        rows: list[RelayPackageRow] = []
        message_table = folder / "聊天记录.xlsx"
        if message_table.exists():
            payload = load_table_rows(message_table)
            for item in payload:
                if not isinstance(item, dict):
                    continue
                message = str(item.get("消息内容", "") or item.get("message", "") or "").strip()
                if not message:
                    continue
                rows.append(
                    RelayPackageRow(
                        sequence=len(rows) + 1,
                        item_type=RelayItemType.TEXT,
                        source_session=source_session,
                        content=message,
                        collected_at=str(item.get("时间", "") or item.get("timestamp", "") or "").strip(),
                    )
                )
        else:
            messages_json = folder / "messages.json"
            if messages_json.exists():
                payload = json.loads(messages_json.read_text(encoding="utf-8"))
                if isinstance(payload, list):
                    message_items = [item for item in payload if isinstance(item, dict)]
                    # 现有导出文件中的 messages.json 保留的是采集原顺序，这里统一回填为“从早到晚”。
                    for item in reversed(message_items):
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
        files_folder = folder / "聊天文件"
        if not files_folder.exists():
            files_folder = folder / "files"
        if files_folder.exists():
            file_items = sorted([path for path in files_folder.iterdir() if path.is_file()], key=lambda item: item.stat().st_mtime)
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
        warning = "" if rows else "文件夹里没有识别到可转发的文本或文件。"
        return RelayCollectionResult(source_session=source_session, rows=rows, warning=warning)

    def _load_relay_package_rows(self, folder: Path, manifest_path: Path) -> RelayCollectionResult:
        if manifest_path.suffix.lower() == ".xlsx":
            items = load_table_rows(manifest_path)
            source_session = self._guess_export_session_name(folder)
            for item in items:
                if not isinstance(item, dict):
                    continue
                candidate = str(item.get("来源会话", "") or item.get("source_session", "") or "").strip()
                if candidate:
                    source_session = candidate
                    break
        else:
            payload = json.loads(manifest_path.read_text(encoding="utf-8"))
            source_session = str(payload.get("source_session", "") or "").strip()
            items = payload.get("items", [])
        rows: list[RelayPackageRow] = []
        if isinstance(items, list):
            for index, item in enumerate(items, start=1):
                if not isinstance(item, dict):
                    continue
                item_type = self._parse_item_type(
                    str(item.get("item_type", "") or item.get("类型", "") or RelayItemType.TEXT.value)
                )
                relative_path = str(item.get("file_path", "") or item.get("相对路径", "") or "").strip()
                absolute_path = ""
                if relative_path:
                    absolute_path = str((folder / relative_path).resolve())
                rows.append(
                    RelayPackageRow(
                        sequence=index,
                        item_type=item_type,
                        source_session=source_session,
                        content=str(item.get("content", "") or item.get("内容预览", "") or "").strip(),
                        file_path=absolute_path,
                        collected_at=str(item.get("collected_at", "") or item.get("采集时间", "") or "").strip(),
                    )
                )
        warning = "" if rows else "发送文件夹中没有识别到可导入的内容。"
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
        package_rows = list(sorted(request.package_rows, key=lambda item: item.sequence))
        if not package_rows:
            raise ValueError("请至少准备一条转发内容。")
        if request.test_only:
            targets = ["文件传输助手"]
        else:
            targets = self._unique_targets([row.downstream_session for row in request.route_rows if row.downstream_session.strip()])
            if not targets:
                raise ValueError("请先填写至少一个收件人会话。")
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
    def keep_latest_file_rows(rows: list[RelayPackageRow]) -> list[RelayPackageRow]:
        latest_by_key: dict[str, tuple[tuple[int, int, int], int]] = {}
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
                    continue
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
    def _placeholder_media_label(message: str) -> str | None:
        normalized = message.strip()
        return MEDIA_PLACEHOLDER_LABELS.get(normalized)

    @staticmethod
    def _sanitize_name(value: str) -> str:
        cleaned = re.sub(r'[\\/:*?"<>|]+', "_", value).strip()
        return cleaned or "send-package"

    @staticmethod
    def _guess_export_session_name(folder: Path) -> str:
        match = re.match(r"^(.*)-\d{8}-\d{6}$", folder.name)
        if match:
            return match.group(1).strip() or folder.name
        return folder.name

    @staticmethod
    def _item_type_label(item_type: RelayItemType) -> str:
        mapping = {
            RelayItemType.TEXT: "文本",
            RelayItemType.FILE: "文件",
            RelayItemType.IMAGE: "图片",
        }
        return mapping[item_type]

    @staticmethod
    def _parse_item_type(value: str) -> RelayItemType:
        normalized = value.strip().lower()
        mapping = {
            "text": RelayItemType.TEXT,
            "文本": RelayItemType.TEXT,
            "file": RelayItemType.FILE,
            "文件": RelayItemType.FILE,
            "image": RelayItemType.IMAGE,
            "图片": RelayItemType.IMAGE,
        }
        return mapping.get(normalized, RelayItemType.TEXT)

    @staticmethod
    def _copy_asset_to_folder(source_path: Path, target_folder: Path, sequence: int) -> Path:
        if not source_path.is_file():
            raise ValueError(f"文件不存在：{source_path}")
        safe_name = source_path.name
        candidate = target_folder / f"{sequence:03d}-{safe_name}"
        counter = 1
        while candidate.exists():
            candidate = target_folder / f"{sequence:03d}-{counter}-{safe_name}"
            counter += 1
        shutil.copy2(source_path, candidate)
        return candidate

    @staticmethod
    def _normalized_file_key(file_path: str) -> str:
        path = Path(file_path)
        stem = re.sub(r"\(\d+\)$", "", path.stem)
        return f"{stem.lower()}{path.suffix.lower()}"

    @staticmethod
    def _file_score(row: RelayPackageRow) -> tuple[int, int, int]:
        path = Path(row.file_path)
        normalized_stem = re.sub(r"\(\d+\)$", "", path.stem)
        canonical_name_score = 1 if path.stem == normalized_stem else 0
        try:
            modified_score = path.stat().st_mtime_ns
        except OSError:
            modified_score = 0
        # On Windows, files created back-to-back can share the same mtime.
        # Prefer the canonical file name first, then the later row order as a stable tiebreaker.
        return (modified_score, canonical_name_score, int(row.sequence))
