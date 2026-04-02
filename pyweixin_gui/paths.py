from __future__ import annotations

from datetime import datetime
import os
from pathlib import Path

APP_NAME = "AutoWeChat"
LEGACY_APP_NAME = "PyWeChatGUI"
ENV_HOME = "AUTOWECHAT_HOME"
LEGACY_ENV_HOME = "PYWEIXIN_GUI_HOME"


def _platform_home() -> Path:
    env_home = os.environ.get(ENV_HOME)
    if env_home:
        return Path(env_home).expanduser().resolve()
    legacy_env_home = os.environ.get(LEGACY_ENV_HOME)
    if legacy_env_home:
        return Path(legacy_env_home).expanduser().resolve()
    appdata = os.environ.get("APPDATA")
    if appdata:
        preferred = Path(appdata) / APP_NAME
        legacy = Path(appdata) / LEGACY_APP_NAME
        if preferred.exists() or not legacy.exists():
            return preferred
        return legacy
    return Path.home() / f".{APP_NAME.lower()}"


def _platform_local_home() -> Path:
    env_home = os.environ.get(ENV_HOME)
    if env_home:
        return Path(env_home).expanduser().resolve()
    legacy_env_home = os.environ.get(LEGACY_ENV_HOME)
    if legacy_env_home:
        return Path(legacy_env_home).expanduser().resolve()
    local_appdata = os.environ.get("LOCALAPPDATA")
    if local_appdata:
        preferred = Path(local_appdata) / APP_NAME
        legacy = Path(local_appdata) / LEGACY_APP_NAME
        if preferred.exists() or not legacy.exists():
            return preferred
        return legacy
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


def create_unique_timestamped_dir(root: str | Path, base_name: str, timestamp: str | None = None) -> Path:
    parent = Path(root).expanduser().resolve()
    parent.mkdir(parents=True, exist_ok=True)
    stamp = timestamp or datetime.now().strftime("%Y%m%d-%H%M%S")
    normalized_base = base_name.strip() or "export"
    suffix = 1

    while True:
        candidate_name = f"{normalized_base}-{stamp}" if suffix == 1 else f"{normalized_base}-{stamp}-{suffix:02d}"
        candidate = parent / candidate_name
        try:
            candidate.mkdir(parents=True, exist_ok=False)
            return candidate
        except FileExistsError:
            suffix += 1
