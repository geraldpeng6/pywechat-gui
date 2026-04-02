from __future__ import annotations

import platform
from pathlib import Path
import time

from .error_handling import UiError, map_exception
from .models import EnvironmentStatus, FileBatchRow, GroupSummaryRow, MessageBatchRow, RelayItemType, RelayPackageRow, RuntimeOptions, SessionSummaryRow
from . import __version__


class UnsupportedPlatformError(RuntimeError):
    """Raised when GUI automation is invoked outside Windows."""


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

    def save_chat_files(self, session_name: str, number: int, target_folder: str, options: RuntimeOptions) -> list[str]:
        pyweixin = self._load_pyweixin()
        return pyweixin.Files.save_chatfiles(
            friend=session_name,
            number=number,
            target_folder=target_folder,
            is_maximize=options.is_maximize,
            close_weixin=options.close_weixin,
        )

    def save_chat_media(self, session_name: str, number: int, target_folder: str, options: RuntimeOptions) -> list[str]:
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
            return []

        items = [item for item in media_list.children() if item.descendants(control_type="Button")]
        if not items:
            chat_history_window.close()
            return []

        items[-1].descendants(control_type="Button")[-1].double_click_input()
        chat_history_window.close()

        desktop = Desktop(backend="uia")
        preview_window = self._find_preview_window(desktop)
        if preview_window is None or not preview_window.exists(timeout=3):
            return []

        saved_paths: list[str] = []
        saved_count = 0
        visited_steps = 0
        max_steps = max(number * 3, number + 5)
        while saved_count < number and visited_steps < max_steps:
            visited_steps += 1
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
