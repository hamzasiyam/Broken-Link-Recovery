"""JSON profile persistence for report and proxy settings.

The GUI workflows use this module to save user-entered settings between runs.
Profiles are intentionally stored as simple JSON files so they can be inspected
and edited manually if needed.
"""

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
    """Small repository object for reading and writing named JSON profiles.

    Args:
        file_path: ``Path`` to the JSON file that stores this profile type.
        defaults: Optional dictionary of default profiles to create when the
            file does not exist.

    Returns:
        A ``ProfileStore`` instance that can load, save, delete, and list
        profiles.
    """

    def __init__(self, file_path: Path, defaults: dict[str, Any] | None = None) -> None:
        """Initialize a profile store for one JSON file.

        Args:
            file_path: ``Path`` object pointing to the JSON file.
            defaults: Optional dictionary used when the JSON file is missing.

        Returns:
            None. The constructor stores configuration on the instance.
        """
        # Keep the path and defaults on the instance so every method uses the
        # same profile file.
        self.file_path = file_path
        self.defaults = defaults or {}

    def load(self) -> dict[str, Any]:
        """Load all profiles from disk, creating defaults if needed.

        Args:
            None.

        Returns:
            Dictionary mapping profile names to profile setting dictionaries.
        """
        # Ensure the profiles directory exists before checking or creating the
        # JSON file.
        ensure_project_directories()

        # If the file is missing, create it from defaults and return those
        # defaults immediately.
        if not self.file_path.exists():
            self._write(dict(self.defaults))
            return dict(self.defaults)

        # If the file exists, parse JSON into a Python dictionary for callers.
        with self.file_path.open("r", encoding="utf-8") as profile_file:
            return json.load(profile_file)

    def save(self, name: str, profile_data: dict[str, Any]) -> None:
        """Save or replace one named profile.

        Args:
            name: Profile name string used as the JSON key.
            profile_data: Dictionary containing the workflow settings to store.

        Returns:
            None. The profile file is written to disk.

        Raises:
            ValueError: If ``name`` is empty.
        """
        # A blank key would be hard to select from the GUI, so reject it before
        # updating the JSON file.
        if not name:
            raise ValueError("Profile name is required.")

        # Load the current set, update one entry, and write the whole JSON file
        # back atomically from the application's perspective.
        profiles = self.load()
        profiles[name] = profile_data
        self._write(profiles)

    def delete(self, name: str) -> bool:
        """Delete one named profile if it exists.

        Args:
            name: Profile name string to remove.

        Returns:
            ``True`` if a profile was removed, otherwise ``False``.
        """
        # Load existing profiles so the delete operation works with the latest
        # file contents.
        profiles = self.load()

        # If the requested profile is absent, tell the caller nothing changed.
        if name not in profiles:
            return False

        # If the requested profile exists, remove it and persist the new file.
        del profiles[name]
        self._write(profiles)
        return True

    def names(self) -> list[str]:
        """Return all profile names for populating GUI comboboxes.

        Args:
            None.

        Returns:
            List of profile name strings.
        """
        # Reuse ``load`` so missing profile files are created before the GUI asks
        # for combobox values.
        return list(self.load().keys())

    def _write(self, profiles: dict[str, Any]) -> None:
        """Write the complete profile dictionary to disk as formatted JSON.

        Args:
            profiles: Dictionary mapping profile names to profile data.

        Returns:
            None. The JSON file is created or overwritten.
        """
        # Ensure the parent folder exists even if this method is called directly
        # by another method before project directories are created.
        self.file_path.parent.mkdir(parents=True, exist_ok=True)

        # Use indentation to keep profile files readable for future engineers
        # and power users who inspect them manually.
        with self.file_path.open("w", encoding="utf-8") as profile_file:
            json.dump(profiles, profile_file, indent=4)


snapshot_profile_store = ProfileStore(SNAPSHOT_PROFILE_FILE, DEFAULT_SNAPSHOT_PROFILES)
proxy_profile_store = ProfileStore(PROXY_PROFILE_FILE, {})
