from pathlib import Path

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Alignment
from openpyxl.utils import get_column_letter

from modules.paths import PROCESSED_DIR
from modules.wayback import convert_to_http, extract_date_from_link, format_date_columns


def save_snapshots_to_excel(snapshots: list[str], filename: str | Path) -> Path:
    if not snapshots:
        raise ValueError("No snapshots were provided.")

    output_path = Path(filename)
    output_path.parent.mkdir(parents=True, exist_ok=True)

    data = [
        (index + 1, snapshot, extract_date_from_link(snapshot))
        for index, snapshot in enumerate(snapshots)
    ]
    df = pd.DataFrame(data, columns=["Capture", "Capture Link", "Date"])
    formatted_dates = df["Date"].apply(format_date_columns).apply(pd.Series)

    df_with_dates = pd.concat([df[["Capture", "Capture Link"]], formatted_dates], axis=1)

    df_http = df.copy()
    df_http["Capture Link"] = df_http["Capture Link"].apply(convert_to_http)
    df_with_dates_http = pd.concat(
        [df_http[["Capture", "Capture Link"]], formatted_dates],
        axis=1,
    )

    summary_df = pd.DataFrame(
        {
            "Total Captures": [len(snapshots)],
            "First Capture Date": [df["Date"].min().strftime("%B %d, %Y %I:%M:%S %p")],
            "Last Capture Date": [df["Date"].max().strftime("%B %d, %Y %I:%M:%S %p")],
        }
    )

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        df_with_dates.to_excel(writer, sheet_name="Snapshots", index=False)
        df_with_dates_http.to_excel(writer, sheet_name="Snapshots_HTTP", index=False)
        summary_df.to_excel(writer, sheet_name="Summary", index=False)

    _format_workbook(output_path)
    return output_path


def snapshot_excel_path(domain: str) -> Path:
    return PROCESSED_DIR / f"{domain}_snapshots.xlsx"


def _format_workbook(filename: Path) -> None:
    workbook = load_workbook(filename)
    for worksheet in workbook.worksheets:
        _autosize_and_align(worksheet)
    workbook.save(filename)


def _autosize_and_align(worksheet) -> None:
    for column in worksheet.columns:
        column_letter = get_column_letter(column[0].column)
        max_length = max(len(str(cell.value or "")) for cell in column)
        worksheet.column_dimensions[column_letter].width = max_length + 2
        for cell in column:
            cell.alignment = Alignment(
                horizontal="center",
                vertical="center",
                wrap_text=True,
            )

    for row in worksheet.iter_rows():
        worksheet.row_dimensions[row[0].row].height = None

