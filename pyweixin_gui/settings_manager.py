from __future__ import annotations

from datetime import datetime
import json
import os
from pathlib import Path
import tempfile

from .models import AppSettings


class SettingsManager:
    def __init__(self, settings_path: Path):
        self.settings_path = settings_path

    def load(self) -> AppSettings:
        if not self.settings_path.exists():
            return AppSettings()
        try:
            data = json.loads(self.settings_path.read_text(encoding="utf-8"))
            if not isinstance(data, dict):
                raise ValueError("settings.json must contain an object")
        except (OSError, UnicodeDecodeError, json.JSONDecodeError, ValueError):
            self._backup_invalid_settings()
            return AppSettings()
        return AppSettings.from_mapping(data)

    def save(self, settings: AppSettings) -> None:
        self.settings_path.parent.mkdir(parents=True, exist_ok=True)
        payload = settings.to_json()
        fd, temp_name = tempfile.mkstemp(
            prefix=f"{self.settings_path.stem}.",
            suffix=".tmp",
            dir=self.settings_path.parent,
        )
        os.close(fd)
        temp_path = Path(temp_name)
        try:
            temp_path.write_text(payload, encoding="utf-8")
            temp_path.replace(self.settings_path)
        finally:
            if temp_path.exists():
                temp_path.unlink(missing_ok=True)

    def _backup_invalid_settings(self) -> None:
        if not self.settings_path.exists():
            return
        timestamp = datetime.now().strftime("%Y%m%d-%H%M%S")
        suffix = self.settings_path.suffix
        stem = self.settings_path.stem
        counter = 1
        while True:
            label = f"{stem}.invalid-{timestamp}" if counter == 1 else f"{stem}.invalid-{timestamp}-{counter:02d}"
            backup_path = self.settings_path.with_name(f"{label}{suffix}")
            if not backup_path.exists():
                break
            counter += 1
        try:
            self.settings_path.replace(backup_path)
        except OSError:
            pass
