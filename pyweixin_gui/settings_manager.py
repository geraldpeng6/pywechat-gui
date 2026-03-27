from __future__ import annotations

import json
from pathlib import Path

from .models import AppSettings


class SettingsManager:
    def __init__(self, settings_path: Path):
        self.settings_path = settings_path

    def load(self) -> AppSettings:
        if not self.settings_path.exists():
            return AppSettings()
        data = json.loads(self.settings_path.read_text(encoding="utf-8"))
        return AppSettings.from_mapping(data)

    def save(self, settings: AppSettings) -> None:
        self.settings_path.write_text(settings.to_json(), encoding="utf-8")
