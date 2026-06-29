from pathlib import Path


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
    """Create runtime directories used by the toolkit."""
    for directory in (
        PROFILE_DIR,
        REPORTS_DIR,
        PROCESSED_DIR,
        RAW_FILES_DIR,
        LOGOS_DIR,
        SCRIPTS_DIR,
    ):
        directory.mkdir(parents=True, exist_ok=True)

