from __future__ import annotations

import platform
from pathlib import Path

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
