from __future__ import annotations

import os
from pathlib import Path


APP_NAME = "PyWeChatGUI"
ENV_HOME = "PYWEIXIN_GUI_HOME"


def _platform_home() -> Path:
    env_home = os.environ.get(ENV_HOME)
    if env_home:
        return Path(env_home).expanduser().resolve()
    appdata = os.environ.get("APPDATA")
    if appdata:
        return Path(appdata) / APP_NAME
    return Path.home() / f".{APP_NAME.lower()}"


def _platform_local_home() -> Path:
    env_home = os.environ.get(ENV_HOME)
    if env_home:
        return Path(env_home).expanduser().resolve()
    local_appdata = os.environ.get("LOCALAPPDATA")
    if local_appdata:
        return Path(local_appdata) / APP_NAME
    return _platform_home()


def ensure_app_dirs() -> dict[str, Path]:
    app_dir = _platform_home()
    local_dir = _platform_local_home()
    logs_dir = local_dir / "logs"
    app_dir.mkdir(parents=True, exist_ok=True)
    local_dir.mkdir(parents=True, exist_ok=True)
    logs_dir.mkdir(parents=True, exist_ok=True)
    return {
        "app_dir": app_dir,
        "local_dir": local_dir,
        "logs_dir": logs_dir,
        "database": app_dir / "app.db",
        "settings": app_dir / "settings.json",
    }
