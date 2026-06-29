# Neverbroken Web Archive Toolkit

Neverbroken is a Python desktop toolkit for working with Wayback Machine website snapshots. It helps export archived capture links to spreadsheets, generate client-ready snapshot reports, and download archived website files for broken-link analysis and recovery work.

## Purpose

The program supports a workflow for finding, organizing, and reviewing archived versions of websites:

- Collect Wayback Machine snapshot URLs for a website.
- Export captures to formatted Excel workbooks.
- Generate Word summary and detailed analysis reports.
- Download archived website contents from snapshot spreadsheets.
- Save reusable report and proxy profiles.

## How to Run

Use `main.py` as the single entry point:

```bash
python main.py
```

The launcher opens a small window where you can choose one of the available workflows.

You can also open a specific workflow directly:

```bash
python main.py --tool snapshot-excel
python main.py --tool snapshot-reports
python main.py --tool archive-downloader
python main.py --tool wget-range
```

List the available workflow IDs:

```bash
python main.py --list-tools
```

The files in `scripts/` are kept only as compatibility wrappers for the older script names. New usage should go through `main.py`.

## Installation

Install Python 3.10 or newer, then install the Python dependencies:

```bash
pip install -r requirements.txt
```

The toolkit also depends on two command-line utilities:

- `waybackpack` for reading Wayback Machine captures.
- `wget` for downloading archived content.

This repository includes `wget.exe` at the project root. The program will use a system `wget` if it is available on `PATH`, otherwise it falls back to the bundled executable.

## Workflows

### Snapshot Excel Exporter

Opens a URL prompt, fetches Wayback Machine captures, and saves a workbook to `reports/processed/<domain>_snapshots.xlsx`. The workbook contains:

- `Snapshots` with original capture links.
- `Snapshots_HTTP` with HTTP-converted links.
- `Summary` with total, first, and last capture dates.

### Snapshot Report Generator

Creates Word documents in `reports/processed/`:

- `<domain>_summary_report.docx`
- `<domain>_detailed_analysis_report.docx`

Profiles for report styling are stored in `profiles/snapshot_analysis_profiles.json`.

### Archive Downloader

Reads a snapshot Excel workbook, lets you choose a sheet, and downloads each capture into an organized folder. Proxy profiles are stored in `profiles/proxy_profiles.json`.

### Wget Range Downloader

Reads capture links from an Excel row range and writes wget log workbooks to `reports/processed/`.

## Project Structure

```text
main.py                         Single application entry point
modules/
  launcher.py                   Main launcher and CLI routing
  paths.py                      Shared project paths and directory setup
  wayback.py                    Wayback URL, domain, snapshot, and date helpers
  excel_reports.py              Excel workbook generation and formatting
  document_reports.py           Word report generation
  downloads.py                  wget checks, downloads, and logging
  proxy.py                      Proxy environment helpers
  profiles.py                   JSON profile persistence
  gui_*.py                      Tkinter workflow windows
scripts/                        Legacy wrappers around the new modules
profiles/                       Saved report and proxy profiles
reports/                        Generated outputs and assets
```

## Technologies Used

- Python
- Tkinter for desktop UI
- pandas for spreadsheet data shaping
- openpyxl for Excel formatting
- python-docx for Word report generation
- Pillow for logo sizing in reports
- waybackpack for Wayback Machine snapshot discovery
- wget for archived file downloads

## Notes

Generated reports and downloads can become large. Keep `reports/processed/`, `downloaded_contents/`, and `downloaded_files/` organized between runs if you are processing many sites.
