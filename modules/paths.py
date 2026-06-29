"""Centralized filesystem paths for Broken Link Recovery Tool.

Every module imports paths from here instead of rebuilding strings manually.
That keeps generated reports, profiles, bundled tools, and download folders
consistent across GUI and command-line workflows.
"""

from pathlib import Path


# Resolve the repository root from this file's location so the app works even
# when launched from a different current working directory.
BASE_DIR = Path(__file__).resolve().parent.parent
PROFILE_DIR = BASE_DIR / "profiles"
REPORTS_DIR = BASE_DIR / "reports"
PROCESSED_DIR = REPORTS_DIR / "processed"
RAW_FILES_DIR = REPORTS_DIR / "raw_files"
LOGOS_DIR = REPORTS_DIR / "all_logos"
SCRIPTS_DIR = BASE_DIR / "scripts"
DOWNLOADED_CONTENTS_DIR = BASE_DIR / "downloaded_contents"
DOWNLOADED_FILES_DIR = BASE_DIR / "downloaded_files"
SNAPSHOT_PROFILE_FILE = PROFILE_DIR / "snapshot_analysis_profiles.json"
PROXY_PROFILE_FILE = PROFILE_DIR / "proxy_profiles.json"
WGET_EXE = BASE_DIR / "wget.exe"


def ensure_project_directories() -> None:
    """Create runtime directories used by the application.

    Args:
        None.

    Returns:
        None. The function creates folders on disk as needed and does not
        return a value.
    """
    # Iterate over every directory the program may write to during normal use.
    for directory in (
        PROFILE_DIR,
        REPORTS_DIR,
        PROCESSED_DIR,
        RAW_FILES_DIR,
        LOGOS_DIR,
        SCRIPTS_DIR,
    ):
        # ``parents=True`` creates missing parent folders; ``exist_ok=True``
        # keeps startup idempotent when folders already exist.
        directory.mkdir(parents=True, exist_ok=True)
