from pathlib import Path

import pandas as pd
from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor

from modules.paths import PROCESSED_DIR
from modules.wayback import extract_date_from_link, split_date_components


def create_summary_report(
    snapshots: list[str],
    domain: str,
    logo_path: str,
    logo_height_percent: str,
    title_color_hex: str,
    heading_color_hex: str,
    output_dir: str | Path = PROCESSED_DIR,
) -> Path:
    if not snapshots:
        raise ValueError("No snapshots were provided.")

    output_path = Path(output_dir) / f"{domain}_summary_report.docx"
    output_path.parent.mkdir(parents=True, exist_ok=True)

    dates = [extract_date_from_link(snapshot) for snapshot in snapshots]
    first_date = min(dates)
    last_date = max(dates)
    first_capture_date, first_capture_time = split_date_components(first_date)
    last_capture_date, last_capture_time = split_date_components(last_date)

    document = Document()
    _add_logo(document, logo_path, logo_height_percent)

    title_run = document.add_heading(f"Snapshot Summary Report for {domain}", 0).runs[0]
    set_font_color(title_run, title_color_hex)

    _add_heading(document, "Introduction", 1, heading_color_hex)
    document.add_paragraph(
        "This report provides a summary of the snapshots captured from the specified URL. "
        "The snapshots represent historical captures of the website, which are the first "
        "step in the process of identifying and fixing broken links. By analyzing these "
        "snapshots, the report shows how the website changed over time."
    )

    _add_heading(document, "Methodology", 1, heading_color_hex)
    document.add_paragraph(
        "The data was collected to provide a timeline of the website's state at various "
        "points in time. These snapshots help identify where broken links may have been "
        "introduced or removed."
    )

    _add_heading(document, "Summary of Captures", 1, heading_color_hex)
    document.add_paragraph(f"Total Captures: {len(snapshots)}")

    _add_heading(document, "First Capture", 2, heading_color_hex)
    document.add_paragraph(f"Date: {first_capture_date}")
    document.add_paragraph(f"Time: {first_capture_time}")

    _add_heading(document, "Last Capture", 2, heading_color_hex)
    document.add_paragraph(f"Date: {last_capture_date}")
    document.add_paragraph(f"Time: {last_capture_time}")

    document.save(output_path)
    return output_path


def create_detailed_analysis_report(
    snapshots: list[str],
    domain: str,
    logo_path: str,
    logo_height_percent: str,
    title_color_hex: str,
    heading_color_hex: str,
    column_color_hex: str,
    output_dir: str | Path = PROCESSED_DIR,
) -> Path:
    if not snapshots:
        raise ValueError("No snapshots were provided.")

    output_path = Path(output_dir) / f"{domain}_detailed_analysis_report.docx"
    output_path.parent.mkdir(parents=True, exist_ok=True)

    data = [
        (index + 1, *split_date_components(extract_date_from_link(snapshot)))
        for index, snapshot in enumerate(snapshots)
    ]
    df = pd.DataFrame(data, columns=["Capture", "Date", "Time"])

    document = Document()
    _add_logo(document, logo_path, logo_height_percent)

    title_run = document.add_heading(f"Detailed Snapshot Analysis for {domain}", 0).runs[0]
    set_font_color(title_run, title_color_hex)

    _add_heading(document, "Introduction", 1, heading_color_hex)
    document.add_paragraph(
        "This document provides a detailed analysis of the snapshots captured from the "
        "specified URL. Each snapshot represents a historical state of the website and "
        "helps guide broken-link analysis."
    )

    _add_heading(document, "Detailed Analysis", 1, heading_color_hex)
    document.add_paragraph(f"Total Captures: {len(snapshots)}")
    document.add_paragraph(
        "The table below lists the snapshots in ascending order, organized by date and time."
    )

    table = document.add_table(rows=1, cols=3)
    header_cells = table.rows[0].cells
    header_cells[0].text = "Capture"
    header_cells[1].text = "Date"
    header_cells[2].text = "Time"

    for cell in header_cells:
        run = cell.paragraphs[0].runs[0]
        run.bold = True
        run.font.size = Pt(12)
        shading = OxmlElement("w:shd")
        shading.set(qn("w:fill"), clean_hex_color(column_color_hex))
        cell._element.get_or_add_tcPr().append(shading)

    for _, row in df.iterrows():
        row_cells = table.add_row().cells
        row_cells[0].text = str(row["Capture"])
        row_cells[1].text = row["Date"]
        row_cells[2].text = row["Time"]

    _add_table_borders(table)

    _add_heading(document, "Next Steps", 1, heading_color_hex)
    _add_bold_paragraph(document, "1. Examination of Snapshots:")
    document.add_paragraph(
        "Each snapshot can be analyzed to check for broken links by reviewing the files "
        "and links in that historical capture."
    )

    _add_bold_paragraph(document, "2. Identification and Fixing of Broken Links:")
    document.add_paragraph("Broken links identified during analysis can be corrected.")

    _add_bold_paragraph(document, "3. Follow-Up Reports:")
    document.add_paragraph(
        "Additional reports can document progress and any findings from the repair process."
    )

    document.save(output_path)
    return output_path


def set_font_color(run, hex_color: str) -> None:
    clean_color = clean_hex_color(hex_color)
    red, green, blue = tuple(int(clean_color[index : index + 2], 16) for index in (0, 2, 4))
    run.font.color.rgb = RGBColor(red, green, blue)


def clean_hex_color(hex_color: str) -> str:
    clean_color = (hex_color or "000000").strip().lstrip("#")
    if len(clean_color) != 6:
        raise ValueError(f"Invalid hex color: {hex_color}")
    int(clean_color, 16)
    return clean_color


def _add_logo(document: Document, logo_path: str, logo_height_percent: str) -> None:
    if not logo_path:
        return

    logo_file = Path(logo_path)
    if not logo_file.exists():
        return

    from PIL import Image

    logo = Image.open(logo_file)
    logo_width, logo_height = logo.size
    height_percent = int(logo_height_percent or 50)
    scaled_height = int(logo_height * height_percent / 100)
    scaled_width = int(logo_width * scaled_height / logo_height)
    document.add_picture(
        str(logo_file),
        width=Inches(scaled_width / 96),
        height=Inches(scaled_height / 96),
    )


def _add_heading(document: Document, text: str, level: int, color_hex: str):
    heading_run = document.add_heading(text, level=level).runs[0]
    set_font_color(heading_run, color_hex)
    return heading_run


def _add_bold_paragraph(document: Document, text: str):
    paragraph = document.add_paragraph()
    run = paragraph.add_run(text)
    run.bold = True
    return paragraph


def _add_table_borders(table) -> None:
    for cell in table._tbl.iter_tcs():
        borders = OxmlElement("w:tcBorders")
        for border_name in ["top", "left", "bottom", "right"]:
            border = OxmlElement(f"w:{border_name}")
            border.set(qn("w:val"), "single")
            border.set(qn("w:sz"), "4")
            border.set(qn("w:space"), "0")
            border.set(qn("w:color"), "auto")
            borders.append(border)
        cell.tcPr.append(borders)

