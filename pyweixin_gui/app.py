from __future__ import annotations

import sys

from .adapter import PyWeixinAdapter
from .logging_utils import configure_logging
from .paths import ensure_app_dirs
from .settings_manager import SettingsManager
from .storage import AppStorage


def main() -> int:
    try:
        from PySide6.QtWidgets import QApplication
    except ImportError as exc:  # pragma: no cover - depends on runtime env
        print("PySide6 未安装，无法启动 GUI。请先安装 GUI 依赖。")
        print(exc)
        return 1

    from .ui.main_window import MainWindow
    from .ui.styles import LIGHT_STYLESHEET

    dirs = ensure_app_dirs()
    logger = configure_logging(dirs["logs_dir"])
    settings_manager = SettingsManager(dirs["settings"])
    settings = settings_manager.load()
    storage = AppStorage(dirs["database"])
    adapter = PyWeixinAdapter()

    app = QApplication(sys.argv)
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
