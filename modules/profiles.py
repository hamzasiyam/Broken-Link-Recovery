import json
from pathlib import Path
from typing import Any

from modules.paths import PROXY_PROFILE_FILE, SNAPSHOT_PROFILE_FILE, ensure_project_directories


DEFAULT_SNAPSHOT_PROFILES = {
    "default": {
        "url": "",
        "logo_path": "",
        "logo_height_percent": "50",
        "title_color_hex": "000000",
        "heading_color_hex": "000000",
        "column_color_hex": "FFFFFF",
    }
}


class ProfileStore:
    def __init__(self, file_path: Path, defaults: dict[str, Any] | None = None) -> None:
        self.file_path = file_path
        self.defaults = defaults or {}

    def load(self) -> dict[str, Any]:
        ensure_project_directories()
        if not self.file_path.exists():
            self._write(dict(self.defaults))
            return dict(self.defaults)

        with self.file_path.open("r", encoding="utf-8") as profile_file:
            return json.load(profile_file)

    def save(self, name: str, profile_data: dict[str, Any]) -> None:
        if not name:
            raise ValueError("Profile name is required.")
        profiles = self.load()
        profiles[name] = profile_data
        self._write(profiles)

    def delete(self, name: str) -> bool:
        profiles = self.load()
        if name not in profiles:
            return False
        del profiles[name]
        self._write(profiles)
        return True

    def names(self) -> list[str]:
        return list(self.load().keys())

    def _write(self, profiles: dict[str, Any]) -> None:
        self.file_path.parent.mkdir(parents=True, exist_ok=True)
        with self.file_path.open("w", encoding="utf-8") as profile_file:
            json.dump(profiles, profile_file, indent=4)


snapshot_profile_store = ProfileStore(SNAPSHOT_PROFILE_FILE, DEFAULT_SNAPSHOT_PROFILES)
proxy_profile_store = ProfileStore(PROXY_PROFILE_FILE, {})

