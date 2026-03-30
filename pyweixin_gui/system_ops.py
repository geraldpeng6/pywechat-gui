from __future__ import annotations

import os
import platform
import subprocess
from pathlib import Path


def open_path(path: str) -> None:
    target = Path(path).expanduser().resolve()
    if not target.exists():
        raise FileNotFoundError(f"路径不存在：{target}")
    system = platform.system()
    if system == "Windows":
        os.startfile(target)  # type: ignore[attr-defined]
        return
    if system == "Darwin":
        subprocess.run(["open", str(target)], check=True)
        return
    subprocess.run(["xdg-open", str(target)], check=True)
