from __future__ import annotations

import sys
from pathlib import Path

if __package__ in {None, ""}:
    workspace_root = Path(__file__).resolve().parent.parent
    if str(workspace_root) not in sys.path:
        sys.path.insert(0, str(workspace_root))

try:
    from .adapter import PyWeixinAdapter
    from .logging_utils import configure_logging
    from .paths import ensure_app_dirs
    from .settings_manager import SettingsManager
    from .storage import AppStorage
except ImportError:
    from pyweixin_gui.adapter import PyWeixinAdapter
    from pyweixin_gui.logging_utils import configure_logging
    from pyweixin_gui.paths import ensure_app_dirs
    from pyweixin_gui.settings_manager import SettingsManager
    from pyweixin_gui.storage import AppStorage


def main() -> int:
    try:
        from PySide6.QtGui import QIcon
        from PySide6.QtWidgets import QApplication
    except ImportError as exc:  # pragma: no cover - depends on runtime env
        print("PySide6 未安装，无法启动 GUI。请先安装 GUI 依赖。")
        print(exc)
        return 1

    try:
        from .ui.main_window import MainWindow
        from .ui.styles import LIGHT_STYLESHEET
    except ImportError:
        from pyweixin_gui.ui.main_window import MainWindow
        from pyweixin_gui.ui.styles import LIGHT_STYLESHEET

    dirs = ensure_app_dirs()
    logger = configure_logging(dirs["logs_dir"])
    settings_manager = SettingsManager(dirs["settings"])
    settings = settings_manager.load()
    storage = AppStorage(dirs["database"])
    adapter = PyWeixinAdapter()

    app = QApplication(sys.argv)
    icon_path = Path(__file__).resolve().parent / "assets" / "autowechat.ico"
    if icon_path.exists():
        app.setWindowIcon(QIcon(str(icon_path)))
    app.setStyleSheet(LIGHT_STYLESHEET)
    window = MainWindow(
        adapter=adapter,
        storage=storage,
        settings_manager=settings_manager,
        settings=settings,
        logger=logger,
    )
    window.show()
    return app.exec()


if __name__ == "__main__":
    raise SystemExit(main())
