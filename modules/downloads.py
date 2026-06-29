import shutil
import subprocess
import time
from pathlib import Path
from typing import Callable, Iterable

import openpyxl
import pandas as pd

from modules.paths import DOWNLOADED_FILES_DIR, PROCESSED_DIR, WGET_EXE


ProgressCallback = Callable[[str], None]


def resolve_wget_executable() -> str:
    system_wget = shutil.which("wget")
    if system_wget:
        return system_wget
    if WGET_EXE.exists():
        return str(WGET_EXE)
    return "wget"


def check_wget(wget_path: str | None = None) -> tuple[bool, str]:
    executable = wget_path or resolve_wget_executable()
    try:
        result = subprocess.run(
            [executable, "--version"],
            capture_output=True,
            text=True,
            check=False,
        )
    except OSError as exc:
        return False, str(exc)

    if result.returncode != 0:
        return False, result.stderr.strip() or "wget is not available."
    first_line = result.stdout.splitlines()[0] if result.stdout else "wget is available."
    return True, first_line


def sanitize_filename(filename: str) -> str:
    return "".join(
        character if character.isalnum() or character in (" ", ".", "_") else "_"
        for character in str(filename)
    )


class WgetDownloader:
    def __init__(self, wget_path: str | None = None) -> None:
        self.wget_path = wget_path or resolve_wget_executable()
        self.current_process: subprocess.Popen | None = None

    def download_archive_snapshot(
        self,
        url: str,
        target_dir: str | Path,
        capture: str,
        date_str: str,
        time_str: str,
        progress_callback: ProgressCallback = print,
    ) -> bool:
        safe_time = str(time_str).replace(":", "_")
        folder_name = sanitize_filename(f"Capture_{capture}_Date_{date_str}_Time_{safe_time}")
        final_target_dir = Path(target_dir) / folder_name
        final_target_dir.mkdir(parents=True, exist_ok=True)
        progress_callback(f"Created directory: {final_target_dir}")

        command = [
            self.wget_path,
            "--mirror",
            "--convert-links",
            "--adjust-extension",
            "--page-requisites",
            "--no-parent",
            str(url),
            "-P",
            str(final_target_dir),
        ]
        progress_callback(f"Running command: {' '.join(command)}")

        try:
            self.current_process = subprocess.Popen(
                command,
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
            )
            if self.current_process.stdout:
                for line in self.current_process.stdout:
                    progress_callback(line.strip())

            self.current_process.wait()
            if self.current_process.returncode != 0:
                progress_callback(f"Error downloading {url}.")
                return False

            progress_callback(f"Successfully downloaded: {url} to {final_target_dir}")
            return True
        except OSError as exc:
            progress_callback(f"Error downloading {url}: {exc}")
            return False
        finally:
            self.current_process = None

    def terminate(self) -> None:
        if self.current_process is not None:
            self.current_process.terminate()
            self.current_process = None


def process_spreadsheet(
    file_path: str | Path,
    sheet_name: str,
    output_dir: str | Path,
    delay_seconds: int,
    downloader: WgetDownloader | None = None,
    progress_callback: ProgressCallback = print,
) -> Path:
    spreadsheet_path = Path(file_path)
    df = pd.read_excel(spreadsheet_path, sheet_name=sheet_name)
    required_columns = [
        "Capture",
        "Capture Link",
        "Month Day, Year",
        "Hour, Minute, Second AM/PM",
    ]
    missing_columns = [column for column in required_columns if column not in df.columns]
    if missing_columns:
        missing = ", ".join(missing_columns)
        raise ValueError(f"The spreadsheet is missing required columns: {missing}")

    domain = spreadsheet_path.name.split("_")[0]
    specific_output_dir = Path(output_dir) / f"{domain}_snapshots"
    specific_output_dir.mkdir(parents=True, exist_ok=True)

    active_downloader = downloader or WgetDownloader()
    for _, row in df.iterrows():
        capture = row["Capture"]
        url = row["Capture Link"]
        date_str = row["Month Day, Year"]
        time_str = row["Hour, Minute, Second AM/PM"]
        progress_callback(f"Downloading URL: {url}")
        success = active_downloader.download_archive_snapshot(
            url,
            specific_output_dir,
            capture,
            date_str,
            time_str,
            progress_callback,
        )
        if not success:
            progress_callback(f"Failed to download: {url}")
        time.sleep(delay_seconds)

    return specific_output_dir


def run_wget_and_log_to_excel(command: Iterable[str], output_excel: str | Path) -> int:
    process = subprocess.Popen(
        list(command),
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=True,
    )
    stdout, stderr = process.communicate()

    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.title = "Wget Log"
    sheet["A1"] = "Timestamp"
    sheet["B1"] = "Event"
    sheet["C1"] = "Details"

    row_index = 2
    from datetime import datetime

    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    for event_name, output in (("Output", stdout), ("Error", stderr)):
        for line in output.splitlines():
            sheet[f"A{row_index}"] = timestamp
            sheet[f"B{row_index}"] = event_name
            sheet[f"C{row_index}"] = line
            row_index += 1

    output_path = Path(output_excel)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    workbook.save(output_path)
    return process.returncode


def download_range_from_excel(
    excel_file: str | Path,
    start_row: int,
    end_row: int,
    download_dir: str | Path = DOWNLOADED_FILES_DIR,
    log_dir: str | Path = PROCESSED_DIR,
) -> list[Path]:
    workbook = openpyxl.load_workbook(excel_file)
    sheet = workbook.active
    wget_path = resolve_wget_executable()
    log_paths: list[Path] = []

    for row in range(start_row, end_row + 1):
        capture_link = sheet.cell(row=row, column=2).value
        if not capture_link:
            continue

        command = [
            wget_path,
            "-r",
            "-l",
            "inf",
            "-np",
            "-nH",
            "--cut-dirs=1",
            "-P",
            str(download_dir),
            "--convert-links",
            "-e",
            "robots=off",
            "-U",
            "Mozilla",
            str(capture_link),
        ]
        output_excel = Path(log_dir) / f"wget_log_{row}.xlsx"
        run_wget_and_log_to_excel(command, output_excel)
        log_paths.append(output_excel)

    return log_paths

