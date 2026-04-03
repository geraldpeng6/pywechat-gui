from __future__ import annotations

from collections import Counter
from datetime import date, datetime, timedelta
import platform
from pathlib import Path
import re
import shutil
import time

from .error_handling import UiError, map_exception
from .models import EnvironmentStatus, FileBatchRow, GroupSummaryRow, MessageBatchRow, RelayCollectMode, RelayItemType, RelayPackageRow, RelayRecentRange, RuntimeOptions, SessionSummaryRow, coerce_relay_recent_range
from . import __version__


class UnsupportedPlatformError(RuntimeError):
    """Raised when GUI automation is invoked outside Windows."""


CHAT_TIMESTAMP_PATTERN = re.compile(r"(?<=\s)(\d{4}年\d{1,2}月\d{1,2}日\s\d{2}:\d{2}|\d{1,2}月\d{1,2}日\s\d{2}:\d{2}|\d{2}:\d{2}|昨天\s\d{2}:\d{2}|星期\w\s\d{2}:\d{2})$")
FILE_TIMESTAMP_PATTERN = re.compile(r"(\d{4}年\d{1,2}月\d{1,2}日|\d{1,2}月\d{1,2}日|昨天|星期\w|\d{1,2}:\d{2})")
WEEKDAY_MAP = {
    "星期一": 0,
    "星期二": 1,
    "星期三": 2,
    "星期四": 3,
    "星期五": 4,
    "星期六": 5,
    "星期天": 6,
    "星期日": 6,
    "Monday": 0,
    "Tuesday": 1,
    "Wednesday": 2,
    "Thursday": 3,
    "Friday": 4,
    "Saturday": 5,
    "Sunday": 6,
}


class PyWeixinAdapter:
    def __init__(self):
        self._pyweixin = None

    def _load_pyweixin(self):
        self._ensure_supported_platform()
        if self._pyweixin is None:
            import pyweixin  # type: ignore

            self._pyweixin = pyweixin
        return self._pyweixin

    @staticmethod
    def _ensure_supported_platform() -> None:
        if platform.system() != "Windows":
            raise UnsupportedPlatformError("该 GUI 仅支持在 Windows 10/11 桌面环境中运行。")

    def inspect_environment(self) -> EnvironmentStatus:
        os_label = f"{platform.system()} {platform.release()}".strip()
        py_version = platform.python_version()
        if platform.system() != "Windows":
            return EnvironmentStatus(
                operating_system=os_label,
                python_version=py_version,
                gui_version=__version__,
                wechat_running=False,
                wechat_path="",
                login_status="当前平台不受支持",
                status_message="当前不是 Windows 环境，无法执行微信桌面自动化。",
                advice=["请在 Windows 10/11 桌面环境中运行该 GUI。"],
            )
        try:
            pyweixin = self._load_pyweixin()
            running = bool(pyweixin.Tools.is_weixin_running())
            wechat_path = ""
            try:
                wechat_path = pyweixin.Tools.where_weixin(copy_to_clipboard=False)
            except Exception:
                wechat_path = ""
            if not running:
                return EnvironmentStatus(
                    operating_system=os_label,
                    python_version=py_version,
                    gui_version=__version__,
                    wechat_running=False,
                    wechat_path=wechat_path,
                    login_status="未启动",
                    status_message="当前未检测到微信进程。",
                    advice=["请先启动微信并保持桌面会话可见。"],
                )
            try:
                pyweixin.Navigator.open_weixin()
                return EnvironmentStatus(
                    operating_system=os_label,
                    python_version=py_version,
                    gui_version=__version__,
                    wechat_running=True,
                    wechat_path=wechat_path,
                    login_status="已登录",
                    status_message="微信连接正常，可以开始执行任务。",
                    advice=["执行期间请勿手工操作微信界面。"],
                )
            except Exception as exc:
                ui_error = map_exception(exc)
                return EnvironmentStatus(
                    operating_system=os_label,
                    python_version=py_version,
                    gui_version=__version__,
                    wechat_running=True,
                    wechat_path=wechat_path,
                    login_status=ui_error.title,
                    status_message=ui_error.message,
                    advice=[ui_error.suggestion],
                )
        except Exception as exc:
            ui_error = map_exception(exc)
            return EnvironmentStatus(
                operating_system=os_label,
                python_version=py_version,
                gui_version=__version__,
                wechat_running=False,
                wechat_path="",
                login_status=ui_error.title,
                status_message=ui_error.message,
                advice=[ui_error.suggestion],
            )

    def send_message(self, row: MessageBatchRow, options: RuntimeOptions) -> None:
        pyweixin = self._load_pyweixin()
        pyweixin.Messages.send_messages_to_friend(
            friend=row.session_name,
            messages=[row.message],
            at_members=row.at_member_list(),
            at_all=row.at_all,
            search_pages=options.search_pages,
            clear=row.clear_before_send,
            send_delay=row.send_delay_sec if row.send_delay_sec is not None else options.send_delay,
            is_maximize=options.is_maximize,
            close_weixin=options.close_weixin,
        )

    def send_file(self, row: FileBatchRow, options: RuntimeOptions) -> None:
        pyweixin = self._load_pyweixin()
        messages = [row.message] if row.with_message and row.message else []
        pyweixin.Files.send_files_to_friend(
            friend=row.session_name,
            files=row.path_list(),
            with_messages=row.with_message,
            messages=messages,
            messages_first=row.message_first,
            send_delay=options.send_delay,
            clear=options.clear,
            is_maximize=options.is_maximize,
            close_weixin=options.close_weixin,
        )

    def dump_chat_history(self, session_name: str, number: int, options: RuntimeOptions) -> tuple[list[str], list[str]]:
        pyweixin = self._load_pyweixin()
        return pyweixin.Messages.dump_chat_history(
            friend=session_name,
            number=number,
            search_pages=options.search_pages,
            is_maximize=options.is_maximize,
            close_weixin=options.close_weixin,
        )

    def dump_recent_chat_history(
        self,
        session_name: str,
        recent_range: RelayRecentRange | str,
        number: int,
        options: RuntimeOptions,
    ) -> tuple[list[str], list[str]]:
        pyweixin = self._load_pyweixin()
        import pyautogui
        from pywinauto import mouse
        from pyweixin.WeChatTools import Navigator, Tools

        normalized_range = coerce_relay_recent_range(recent_range)
        if not isinstance(normalized_range, RelayRecentRange):
            raise ValueError("请选择有效的时间范围")

        chat_history_window = Navigator.open_chat_history(
            friend=session_name,
            search_pages=options.search_pages,
            is_maximize=options.is_maximize,
            close_weixin=options.close_weixin,
        )
        chat_history_list = chat_history_window.child_window(title="聊天记录", control_type="List")
        if not chat_history_list.exists(timeout=0.3):
            chat_history_window.close()
            return [], []

        Tools.activate_chatHistoryList(chat_history_list)
        items = chat_history_list.children(control_type="ListItem")
        if not items:
            chat_history_window.close()
            return [], []

        last_item = items[-1]
        rect = last_item.rectangle()
        mouse.click(coords=(rect.right - 30, rect.bottom - 20))
        now = datetime.now()
        messages: list[str] = []
        timestamps: list[str] = []
        previous_runtime_id: tuple[int, ...] | None = None

        while len(messages) < number:
            selected = [item for item in chat_history_list.children(control_type="ListItem") if item.is_selected() or item.has_keyboard_focus()]
            current_item = selected[0] if selected else last_item
            runtime_id = tuple(current_item.element_info.runtime_id)
            if previous_runtime_id is not None and runtime_id == previous_runtime_id:
                break
            previous_runtime_id = runtime_id

            raw_text = current_item.window_text()
            timestamp_text = self._extract_chat_timestamp(raw_text)
            if timestamp_text:
                if self._recent_label_in_range(timestamp_text, normalized_range, now):
                    messages.append(self._strip_chat_timestamp(raw_text))
                    timestamps.append(timestamp_text)
                else:
                    break
            pyautogui.press("up", presses=2, _pause=False)

        chat_history_list.type_keys("{HOME}")
        chat_history_window.close()
        return messages, timestamps

    def dump_chat_history_items(
        self,
        session_name: str,
        number: int,
        options: RuntimeOptions,
        recent_range: RelayRecentRange | str | None = None,
    ) -> list[dict[str, str]]:
        self._load_pyweixin()
        import pyautogui
        from pywinauto import mouse
        from pyweixin.WeChatTools import Navigator, Tools

        normalized_range = None if recent_range is None else coerce_relay_recent_range(recent_range)
        if normalized_range is not None and not isinstance(normalized_range, RelayRecentRange):
            raise ValueError("请选择有效的时间范围")

        chat_history_window = Navigator.open_chat_history(
            friend=session_name,
            search_pages=options.search_pages,
            is_maximize=options.is_maximize,
            close_weixin=options.close_weixin,
        )
        chat_history_list = chat_history_window.child_window(title="聊天记录", control_type="List")
        if not chat_history_list.exists(timeout=0.3):
            chat_history_window.close()
            return []

        Tools.activate_chatHistoryList(chat_history_list)
        items = chat_history_list.children(control_type="ListItem")
        if not items:
            chat_history_window.close()
            return []

        last_item = items[-1]
        rect = last_item.rectangle()
        mouse.click(coords=(rect.right - 30, rect.bottom - 20))

        now = datetime.now()
        results: list[dict[str, str]] = []
        previous_runtime_id: tuple[int, ...] | None = None
        while len(results) < number:
            selected = [item for item in chat_history_list.children(control_type="ListItem") if item.is_selected() or item.has_keyboard_focus()]
            current_item = selected[0] if selected else last_item
            runtime_id = tuple(current_item.element_info.runtime_id)
            if previous_runtime_id is not None and runtime_id == previous_runtime_id:
                break
            previous_runtime_id = runtime_id

            parsed = self._parse_chat_history_item(current_item)
            if parsed["timestamp"] and normalized_range is not None and not self._recent_label_in_range(parsed["timestamp"], normalized_range, now):
                break
            results.append(parsed)
            pyautogui.press("up", presses=2, _pause=False)

        chat_history_list.type_keys("{HOME}")
        chat_history_window.close()
        return results

    def save_chat_files(self, session_name: str, number: int, target_folder: str, options: RuntimeOptions) -> list[str]:
        pyweixin = self._load_pyweixin()
        return pyweixin.Files.save_chatfiles(
            friend=session_name,
            number=number,
            target_folder=target_folder,
            is_maximize=options.is_maximize,
            close_weixin=options.close_weixin,
        )

    def save_recent_chat_files(
        self,
        session_name: str,
        recent_range: RelayRecentRange | str,
        number: int,
        target_folder: str,
        options: RuntimeOptions,
    ) -> list[str]:
        pyweixin = self._load_pyweixin()
        import pyautogui
        from pywinauto import mouse
        from pyweixin.WeChatTools import Navigator

        normalized_range = coerce_relay_recent_range(recent_range)
        if not isinstance(normalized_range, RelayRecentRange):
            raise ValueError("请选择有效的时间范围")

        target_dir = Path(target_folder).expanduser().resolve()
        target_dir.mkdir(parents=True, exist_ok=True)
        chatfile_root = Path(pyweixin.Tools.where_chatfile_folder())

        chat_history_window = Navigator.open_chat_history(
            friend=session_name,
            TabItem="文件",
            search_pages=options.search_pages,
            is_maximize=options.is_maximize,
            close_weixin=options.close_weixin,
        )
        file_list = self._find_chat_file_list(chat_history_window)
        if file_list is None or not file_list.exists(timeout=0.3):
            chat_history_window.close()
            return []

        rect = file_list.rectangle()
        mouse.click(coords=(rect.right - 8, rect.top + 5))
        pyautogui.press("home")

        items = file_list.children(control_type="ListItem")
        if not items:
            chat_history_window.close()
            return []
        first_rect = items[0].rectangle()
        mouse.click(coords=(first_rect.right - 20, first_rect.bottom - 5))

        now = datetime.now()
        source_paths: list[Path] = []
        previous_runtime_id: tuple[int, ...] | None = None
        while len(source_paths) < number:
            selected = [item for item in file_list.children(control_type="ListItem") if item.is_selected() or item.has_keyboard_focus()]
            if not selected:
                break
            current_item = selected[0]
            runtime_id = tuple(current_item.element_info.runtime_id)
            if previous_runtime_id is not None and runtime_id == previous_runtime_id:
                break
            previous_runtime_id = runtime_id

            timestamp_text = self._extract_file_timestamp(current_item)
            if timestamp_text and self._recent_label_in_range(timestamp_text, normalized_range, now):
                source_path = self._resolve_chat_file_source_path(chatfile_root, current_item, timestamp_text, now)
                if source_path is not None and source_path.exists():
                    source_paths.append(source_path)
            elif timestamp_text:
                break
            pyautogui.press("down", _pause=False)

        chat_history_window.close()
        copied_paths = self._copy_recent_chat_files(source_paths, target_dir)
        return [str(path) for path in copied_paths]

    def list_chat_file_items(
        self,
        session_name: str,
        number: int,
        options: RuntimeOptions,
        recent_range: RelayRecentRange | str | None = None,
    ) -> list[dict[str, str]]:
        pyweixin = self._load_pyweixin()
        import pyautogui
        from pywinauto import mouse
        from pyweixin.WeChatTools import Navigator

        normalized_range = None if recent_range is None else coerce_relay_recent_range(recent_range)
        if normalized_range is not None and not isinstance(normalized_range, RelayRecentRange):
            raise ValueError("请选择有效的时间范围")

        chatfile_root = Path(pyweixin.Tools.where_chatfile_folder())
        chat_history_window = Navigator.open_chat_history(
            friend=session_name,
            TabItem="文件",
            search_pages=options.search_pages,
            is_maximize=options.is_maximize,
            close_weixin=options.close_weixin,
        )
        file_list = self._find_chat_file_list(chat_history_window)
        if file_list is None or not file_list.exists(timeout=0.3):
            chat_history_window.close()
            return []

        rect = file_list.rectangle()
        mouse.click(coords=(rect.right - 8, rect.top + 5))
        pyautogui.press("home")

        items = file_list.children(control_type="ListItem")
        if not items:
            chat_history_window.close()
            return []
        first_rect = items[0].rectangle()
        mouse.click(coords=(first_rect.right - 20, first_rect.bottom - 5))

        now = datetime.now()
        results: list[dict[str, str]] = []
        previous_runtime_id: tuple[int, ...] | None = None
        while len(results) < number:
            selected = [item for item in file_list.children(control_type="ListItem") if item.is_selected() or item.has_keyboard_focus()]
            if not selected:
                break
            current_item = selected[0]
            runtime_id = tuple(current_item.element_info.runtime_id)
            if previous_runtime_id is not None and runtime_id == previous_runtime_id:
                break
            previous_runtime_id = runtime_id

            timestamp_text = self._extract_file_timestamp(current_item)
            if timestamp_text and normalized_range is not None and not self._recent_label_in_range(timestamp_text, normalized_range, now):
                break
            source_path = self._resolve_chat_file_source_path(chatfile_root, current_item, timestamp_text, now)
            if source_path is not None and source_path.exists():
                results.append(
                    {
                        "sender": self._extract_file_sender(current_item),
                        "timestamp": timestamp_text,
                        "source_path": str(source_path),
                        "name": source_path.name,
                    }
                )
            pyautogui.press("down", _pause=False)

        chat_history_window.close()
        return results

    def save_chat_media(self, session_name: str, number: int, target_folder: str, options: RuntimeOptions) -> list[str]:
        preview_window, target_dir = self._open_chat_media_preview(
            session_name=session_name,
            target_folder=target_folder,
            options=options,
        )
        if preview_window is None:
            return []
        return self._save_media_from_preview(
            preview_window=preview_window,
            target_dir=target_dir,
            session_name=session_name,
            save_limit=number,
        )

    def save_recent_chat_media(
        self,
        session_name: str,
        recent_range: RelayRecentRange | str,
        number: int,
        target_folder: str,
        options: RuntimeOptions,
    ) -> list[str]:
        from pywinauto import Desktop
        import pyautogui

        from pyweixin.WeChatTools import Navigator

        normalized_range = coerce_relay_recent_range(recent_range)
        if not isinstance(normalized_range, RelayRecentRange):
            raise ValueError("请选择有效的时间范围")

        target_dir = Path(target_folder).expanduser().resolve()
        target_dir.mkdir(parents=True, exist_ok=True)
        pyautogui.FAILSAFE = False

        chat_history_window = Navigator.open_chat_history(
            friend=session_name,
            TabItem="图片与视频",
            search_pages=options.search_pages,
            is_maximize=options.is_maximize,
            close_weixin=options.close_weixin,
        )
        media_list = self._find_media_history_list(chat_history_window)
        if media_list is None or not media_list.exists(timeout=0.5):
            chat_history_window.close()
            return []

        items = [item for item in media_list.children() if item.descendants(control_type="Button")]
        if not items:
            chat_history_window.close()
            return []

        visit_limit = self._count_recent_media_visit_limit(items, normalized_range, datetime.now(), number)
        if visit_limit <= 0:
            chat_history_window.close()
            return []

        items[-1].descendants(control_type="Button")[-1].double_click_input()
        chat_history_window.close()

        desktop = Desktop(backend="uia")
        preview_window = self._find_preview_window(desktop)
        if preview_window is None or not preview_window.exists(timeout=3):
            return []
        return self._save_media_from_preview(
            preview_window=preview_window,
            target_dir=target_dir,
            session_name=session_name,
            save_limit=number,
            visit_limit=visit_limit,
        )

    def _open_chat_media_preview(self, session_name: str, target_folder: str, options: RuntimeOptions):
        from pywinauto import Desktop
        import pyautogui

        from pyweixin.WeChatTools import Navigator

        target_dir = Path(target_folder).expanduser().resolve()
        target_dir.mkdir(parents=True, exist_ok=True)
        pyautogui.FAILSAFE = False

        chat_history_window = Navigator.open_chat_history(
            friend=session_name,
            TabItem="图片与视频",
            search_pages=options.search_pages,
            is_maximize=options.is_maximize,
            close_weixin=options.close_weixin,
        )
        media_list = self._find_media_history_list(chat_history_window)
        if media_list is None or not media_list.exists(timeout=0.5):
            chat_history_window.close()
            return None, target_dir

        items = [item for item in media_list.children() if item.descendants(control_type="Button")]
        if not items:
            chat_history_window.close()
            return None, target_dir

        items[-1].descendants(control_type="Button")[-1].double_click_input()
        chat_history_window.close()

        desktop = Desktop(backend="uia")
        preview_window = self._find_preview_window(desktop)
        if preview_window is None or not preview_window.exists(timeout=3):
            return None, target_dir
        return preview_window, target_dir

    def _save_media_from_preview(
        self,
        preview_window,
        target_dir: Path,
        session_name: str,
        save_limit: int,
        visit_limit: int | None = None,
    ) -> list[str]:
        import pyautogui

        saved_paths: list[str] = []
        saved_count = 0
        visited_items = 0
        max_visits = visit_limit if visit_limit is not None else max(save_limit * 3, save_limit + 5)
        while saved_count < save_limit and visited_items < max_visits:
            visited_items += 1
            media_kind = self._detect_preview_media_kind(preview_window)
            if media_kind is None:
                self._move_preview_to_previous(pyautogui)
                if self._is_preview_at_first(preview_window):
                    break
                continue

            saved_path: str | None = None
            if media_kind == "image":
                saved_path = self._capture_preview_image(preview_window, target_dir, session_name, saved_count + 1)
            elif media_kind == "video":
                saved_path = self._save_preview_video(preview_window, target_dir)
            if saved_path:
                saved_paths.append(saved_path)
                saved_count += 1

            self._move_preview_to_previous(pyautogui)
            if self._is_preview_at_first(preview_window):
                break

        if preview_window.exists(timeout=0.3):
            preview_window.close()
        return saved_paths

    def export_recent_files(self, target_folder: str, options: RuntimeOptions) -> list[str]:
        pyweixin = self._load_pyweixin()
        return pyweixin.Files.export_recent_files(
            target_folder=target_folder,
            is_maximize=options.is_maximize,
            close_weixin=options.close_weixin,
        )

    def export_wxfiles(self, year: str, month: str | None, target_folder: str) -> list[str]:
        pyweixin = self._load_pyweixin()
        return pyweixin.Files.export_wxfiles(
            year=year,
            month=month,
            target_folder=target_folder,
        )

    def export_videos(self, year: str, month: str | None, target_folder: str) -> list[str]:
        pyweixin = self._load_pyweixin()
        try:
            return pyweixin.Files.export_videos(
                year=year,
                month=month,
                target_folder=target_folder,
            )
        except UnboundLocalError as exc:
            if "exported_videos" not in str(exc):
                raise
            folder = Path(target_folder)
            if not folder.exists():
                return []
            return [str(path) for path in sorted(folder.iterdir()) if path.is_file()]

    def dump_sessions(self, chatted_only: bool, options: RuntimeOptions) -> list[SessionSummaryRow]:
        pyweixin = self._load_pyweixin()
        sessions = pyweixin.Messages.dump_sessions(
            chat_only=chatted_only,
            is_maximize=options.is_maximize,
            close_weixin=options.close_weixin,
        )
        return [
            SessionSummaryRow(
                session_name=str(session[0] or "").strip(),
                last_time=str(session[1] or "").strip() if len(session) > 1 else "",
                last_message=str(session[2] or "").strip() if len(session) > 2 else "",
            )
            for session in sessions
            if session and str(session[0] or "").strip()
        ]

    def dump_groups(self, options: RuntimeOptions) -> list[GroupSummaryRow]:
        pyweixin = self._load_pyweixin()
        groups = pyweixin.Contacts.get_groups_info(
            is_maximize=options.is_maximize,
            close_weixin=options.close_weixin,
        )
        return [
            GroupSummaryRow(
                group_name=str(group or "").strip(),
            )
            for group in groups
            if str(group or "").strip()
        ]

    def validate_session(self, session_name: str, options: RuntimeOptions) -> None:
        pyweixin = self._load_pyweixin()
        pyweixin.Navigator.open_dialog_window(
            friend=session_name,
            is_maximize=options.is_maximize,
            search_pages=options.search_pages,
        )

    def send_relay_item(self, target_session: str, item: RelayPackageRow, options: RuntimeOptions) -> None:
        if item.item_type is RelayItemType.TEXT:
            row = MessageBatchRow(
                session_name=target_session,
                message=item.content,
                clear_before_send=options.clear,
            )
            self.send_message(row, options)
            return
        row = FileBatchRow(
            session_name=target_session,
            file_paths=item.file_path,
            with_message=False,
        )
        self.send_file(row, options)

    @staticmethod
    def map_runtime_exception(exc: BaseException) -> UiError:
        return map_exception(exc)

    @staticmethod
    def _find_media_history_list(chat_history_window):
        selectors = [
            {"title": "照片和视频", "control_type": "List"},
            {"title": "图片与视频", "control_type": "List"},
            {"title": "Photos & Videos", "control_type": "List"},
            {"title": "圖片與影片", "control_type": "List"},
        ]
        for selector in selectors:
            widget = chat_history_window.child_window(**selector)
            if widget.exists(timeout=0.2):
                return widget
        return None

    @staticmethod
    def _find_chat_file_list(chat_history_window):
        selectors = [
            {"control_type": "List", "auto_id": "file_list"},
            {"title": "文件", "control_type": "List"},
        ]
        for selector in selectors:
            widget = chat_history_window.child_window(**selector)
            if widget.exists(timeout=0.2):
                return widget
        return None

    @staticmethod
    def _find_preview_window(desktop):
        selectors = [
            {"control_type": "Window", "class_name": "mmui::PreviewWindow"},
            {"control_type": "Window", "class_name": "ImagePreviewWnd", "framework_id": "Win32"},
        ]
        for selector in selectors:
            widget = desktop.window(**selector)
            if widget.exists(timeout=0.5):
                return widget
        return None

    @staticmethod
    def _detect_preview_media_kind(preview_window) -> str | None:
        image_expired = preview_window.child_window(
            title_re="图片过期或已被清理|Image expired or deleted|圖片過期或已被刪除",
            control_type="Text",
        )
        video_expired = preview_window.child_window(
            title_re="视频过期或已被删除|Video expired or deleted\\.?|影片已逾期或已被刪除",
            control_type="Text",
        )
        if image_expired.exists(timeout=0.2) or video_expired.exists(timeout=0.2):
            return None
        rotate_button = preview_window.child_window(title_re="旋转|Rotate|旋轉", control_type="Button")
        if rotate_button.exists(timeout=0.2):
            return "image"
        return "video"

    @staticmethod
    def _move_preview_to_previous(pyautogui_module) -> None:
        pyautogui_module.press("left", _pause=False)
        time.sleep(0.2)

    @staticmethod
    def _is_preview_at_first(preview_window) -> bool:
        earliest_text = preview_window.child_window(
            title_re="已是第一张|This is the first one|已是第一張",
            control_type="Text",
        )
        return earliest_text.exists(timeout=0.2)

    @staticmethod
    def _capture_preview_image(preview_window, target_dir: Path, session_name: str, index: int) -> str | None:
        candidates = preview_window.descendants(control_type="Button", title="")
        if not candidates:
            return None
        image_button = max(
            candidates,
            key=lambda item: max(1, item.rectangle().width()) * max(1, item.rectangle().height()),
        )
        image_area = image_button.parent()
        if image_area is None:
            return None
        timestamp = time.strftime("%Y%m%d-%H%M%S")
        safe_name = "".join(char if char.isalnum() or char in {"-", "_"} else "_" for char in session_name).strip("_") or "chat"
        path = target_dir / f"{safe_name}-图片-{timestamp}-{index:03d}.png"
        image = image_area.capture_as_image()
        image.save(path)
        return str(path)

    @staticmethod
    def _save_preview_video(preview_window, target_dir: Path) -> str | None:
        video_player = preview_window.child_window(title="player video", control_type="Pane")
        save_button = preview_window.child_window(title_re="另存为|Save as|另存為", control_type="Button")
        if video_player.exists(timeout=0.5):
            try:
                video_player.wait(wait_for="ready", timeout=15, retry_interval=0.3)
            except Exception:
                pass
        if save_button.exists(timeout=1):
            try:
                save_button.wait(wait_for="enabled", timeout=3, retry_interval=0.3)
            except Exception:
                pass
        if not save_button.exists(timeout=0.5):
            return None
        before_names = {path.name for path in target_dir.iterdir() if path.is_file()}
        save_button.click_input()
        return PyWeixinAdapter._save_preview_to_folder(target_dir, before_names)

    @staticmethod
    def _extract_chat_timestamp(raw_text: str) -> str:
        match = CHAT_TIMESTAMP_PATTERN.search(str(raw_text or "").strip())
        return match.group(1).strip() if match else ""

    @staticmethod
    def _strip_chat_timestamp(raw_text: str) -> str:
        text = str(raw_text or "").strip()
        return CHAT_TIMESTAMP_PATTERN.sub("", text).strip()

    @staticmethod
    def _extract_file_timestamp(item) -> str:
        texts = item.descendants(control_type="Text")
        if len(texts) < 2:
            return ""
        candidate = str(texts[-2].window_text() or "").strip()
        match = FILE_TIMESTAMP_PATTERN.search(candidate)
        return match.group(1).strip() if match else candidate

    @staticmethod
    def _extract_media_timestamp(item) -> str:
        candidates = [str(item.window_text() or "").strip()]
        candidates.extend(str(text.window_text() or "").strip() for text in item.descendants(control_type="Text"))
        for candidate in reversed(candidates):
            if not candidate:
                continue
            match = FILE_TIMESTAMP_PATTERN.search(candidate)
            if match:
                return match.group(1).strip()
        return ""

    @staticmethod
    def _count_recent_media_visit_limit(items: list, recent_range: RelayRecentRange, now: datetime, save_limit: int) -> int:
        visit_limit = 0
        matched_items = 0
        for item in reversed(items):
            timestamp_text = PyWeixinAdapter._extract_media_timestamp(item)
            if timestamp_text:
                if not PyWeixinAdapter._recent_label_in_range(timestamp_text, recent_range, now):
                    break
                matched_items += 1
            visit_limit += 1
            if matched_items >= save_limit:
                break
        return visit_limit if matched_items > 0 else 0

    @staticmethod
    def _extract_file_sender(item) -> str:
        filename = str(item.window_text() or "").strip()
        texts = [str(text.window_text() or "").strip() for text in item.descendants(control_type="Text")]
        size_pattern = re.compile(r"^\d+(\.\d+)?\s*(B|KB|MB|GB)$", re.IGNORECASE)
        for text in texts:
            if not text or text == filename:
                continue
            if FILE_TIMESTAMP_PATTERN.search(text):
                continue
            if size_pattern.search(text):
                continue
            return text
        return ""

    @staticmethod
    def _resolve_chat_file_source_path(chatfile_root: Path, item, timestamp_text: str, now: datetime) -> Path | None:
        file_name = str(item.window_text() or "").strip()
        if not file_name:
            return None
        parsed_date = PyWeixinAdapter._parse_recent_label_date(timestamp_text, now)
        if parsed_date is None:
            return None
        month_folder = f"{parsed_date.year}-{parsed_date.month:02d}"
        return chatfile_root / month_folder / file_name

    @staticmethod
    def _copy_recent_chat_files(source_paths: list[Path], target_dir: Path) -> list[Path]:
        copied_paths: list[Path] = []
        counts = Counter(source_paths)
        handled_sources: set[Path] = set()
        for source_path in source_paths:
            if source_path in handled_sources:
                continue
            handled_sources.add(source_path)
            if counts[source_path] > 1:
                candidates = PyWeixinAdapter._duplicate_file_candidates(source_path)
            else:
                candidates = [source_path]
            for candidate in candidates:
                if candidate.exists():
                    copied_paths.append(PyWeixinAdapter._copy_file_preserving_name(candidate, target_dir))
        return copied_paths

    @staticmethod
    def _duplicate_file_candidates(source_path: Path) -> list[Path]:
        folder = source_path.parent
        base_name = source_path.name
        stem = source_path.stem
        suffix = source_path.suffix
        pattern = re.compile(rf"^{re.escape(stem)}\(\d+\){re.escape(suffix)}$")
        candidates = [source_path]
        for candidate in sorted(folder.iterdir()):
            if candidate.name == source_path.name:
                continue
            if pattern.match(candidate.name):
                candidates.append(candidate)
        return candidates

    @staticmethod
    def _copy_file_preserving_name(source_path: Path, target_dir: Path) -> Path:
        candidate = target_dir / source_path.name
        if not candidate.exists():
            shutil.copy2(source_path, candidate)
            return candidate
        counter = 1
        while True:
            renamed = target_dir / f"{source_path.stem}({counter}){source_path.suffix}"
            if not renamed.exists():
                shutil.copy2(source_path, renamed)
                return renamed
            counter += 1

    @staticmethod
    def _recent_label_in_range(label: str, recent_range: RelayRecentRange, now: datetime) -> bool:
        parsed_date = PyWeixinAdapter._parse_recent_label_date(label, now)
        if parsed_date is None:
            return False
        current_date = now.date()
        if recent_range is RelayRecentRange.TODAY:
            return parsed_date == current_date
        if recent_range is RelayRecentRange.YESTERDAY:
            return parsed_date == current_date - timedelta(days=1)
        if recent_range is RelayRecentRange.WEEK:
            week_start = current_date - timedelta(days=current_date.weekday())
            return week_start <= parsed_date <= current_date
        if recent_range is RelayRecentRange.MONTH:
            return parsed_date.year == current_date.year and parsed_date.month == current_date.month and parsed_date <= current_date
        return False

    @staticmethod
    def _parse_recent_label_date(label: str, now: datetime) -> date | None:
        text = str(label or "").strip()
        if not text:
            return None
        if re.fullmatch(r"\d{1,2}:\d{2}", text):
            return now.date()
        if re.fullmatch(r"昨天(?:\s+\d{1,2}:\d{2})?", text) or re.fullmatch(r"Yesterday(?:\s+\d{1,2}:\d{2})?", text):
            return now.date() - timedelta(days=1)
        weekday_match = re.fullmatch(r"(星期[一二三四五六天日]|Monday|Tuesday|Wednesday|Thursday|Friday|Saturday|Sunday)(?:\s+\d{1,2}:\d{2})?", text)
        if weekday_match:
            weekday_name = weekday_match.group(1)
            target_weekday = WEEKDAY_MAP.get(weekday_name)
            if target_weekday is None:
                return None
            week_start = now.date() - timedelta(days=now.weekday())
            candidate = week_start + timedelta(days=target_weekday)
            if candidate > now.date():
                candidate -= timedelta(days=7)
            return candidate
        full_match = re.fullmatch(r"(\d{4})年(\d{1,2})月(\d{1,2})日(?:\s+\d{1,2}:\d{2})?", text)
        if full_match:
            year, month, day = map(int, full_match.groups())
            return date(year, month, day)
        month_day_match = re.fullmatch(r"(\d{1,2})月(\d{1,2})日(?:\s+\d{1,2}:\d{2})?", text)
        if month_day_match:
            month, day = map(int, month_day_match.groups())
            year = now.year
            candidate = date(year, month, day)
            if candidate > now.date():
                candidate = date(year - 1, month, day)
            return candidate
        return None

    @staticmethod
    def _parse_chat_history_item(item) -> dict[str, str]:
        texts = [str(text.window_text() or "").strip() for text in item.descendants(control_type="Text") if str(text.window_text() or "").strip()]
        item_text = str(item.window_text() or "").strip()
        sender = texts[0] if texts else ""
        timestamp = texts[1] if len(texts) > 1 else PyWeixinAdapter._extract_chat_timestamp(item_text)
        special_map = {
            "[图片]": "图片",
            "图片": "图片",
            "[视频]": "视频",
            "视频": "视频",
            "[动画表情]": "动画表情",
            "动画表情": "动画表情",
        }
        content = ""
        if item_text in special_map:
            content = special_map[item_text]
        elif len(texts) >= 3:
            content = texts[2]
        if not content:
            content = PyWeixinAdapter._strip_chat_timestamp(item_text)
        return {"sender": sender, "timestamp": timestamp, "content": content}

    @staticmethod
    def _save_preview_to_folder(target_dir: Path, before_names: set[str]) -> str | None:
        from pywinauto import Desktop
        import pyautogui

        from pyweixin.WinSettings import SystemSettings

        SystemSettings.copy_text_to_clipboard(str(target_dir))
        desktop = Desktop(backend="uia")
        save_window = desktop.window(
            title_re="另存为|Save as|另存為",
            control_type="Window",
            framework_id="Win32",
            top_level_only=False,
        )
        if not save_window.exists(timeout=5):
            return None
        save_window.set_focus()
        time.sleep(0.2)
        pyautogui.hotkey("alt", "d", _pause=False)
        time.sleep(0.2)
        pyautogui.hotkey("ctrl", "a", _pause=False)
        pyautogui.hotkey("ctrl", "v", _pause=False)
        pyautogui.press("enter", _pause=False)
        time.sleep(0.3)
        pyautogui.hotkey("alt", "s", _pause=False)

        deadline = time.time() + 15
        while time.time() < deadline:
            paths = [path for path in target_dir.iterdir() if path.is_file()]
            new_paths = [path for path in paths if path.name not in before_names]
            if new_paths:
                newest = max(new_paths, key=lambda item: item.stat().st_mtime_ns)
                return str(newest)
            time.sleep(0.3)
        return None
