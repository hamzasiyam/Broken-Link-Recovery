import argparse
import importlib
import tkinter as tk
from tkinter import messagebox

from modules.paths import ensure_project_directories


TOOL_DEFINITIONS = {
    "snapshot-excel": {
        "label": "Snapshot Excel Exporter",
        "module": "modules.gui_snapshot_exporter",
        "description": "Export Wayback Machine capture links into an Excel workbook.",
    },
    "snapshot-reports": {
        "label": "Snapshot Report Generator",
        "module": "modules.gui_report_generator",
        "description": "Generate Word summary and detailed analysis reports.",
    },
    "archive-downloader": {
        "label": "Archive Downloader",
        "module": "modules.gui_archive_downloader",
        "description": "Download archived website files from a snapshot spreadsheet.",
    },
    "wget-range": {
        "label": "Wget Range Downloader",
        "module": "modules.gui_wget_range",
        "description": "Download a row range from an Excel workbook and log wget output.",
    },
}


def launch_tool(tool_id: str, parent: tk.Tk | tk.Toplevel | None = None):
    tool = TOOL_DEFINITIONS[tool_id]
    module = importlib.import_module(tool["module"])
    return module.launch(parent)


def launch_tool_selector() -> None:
    ensure_project_directories()

    root = tk.Tk()
    root.title("Neverbroken Web Archive Toolkit")
    root.geometry("460x310")

    frame = tk.Frame(root, padx=16, pady=16)
    frame.pack(fill="both", expand=True)

    title = tk.Label(
        frame,
        text="Neverbroken Web Archive Toolkit",
        font=("Arial", 14, "bold"),
    )
    title.pack(anchor="w", pady=(0, 8))

    subtitle = tk.Label(
        frame,
        text="Choose a workflow to open.",
        font=("Arial", 10),
    )
    subtitle.pack(anchor="w", pady=(0, 12))

    for tool_id, tool in TOOL_DEFINITIONS.items():
        button = tk.Button(
            frame,
            text=tool["label"],
            width=34,
            command=lambda selected_tool=tool_id: _open_child_tool(root, selected_tool),
        )
        button.pack(anchor="w", pady=4)

    tk.Button(frame, text="Close", width=34, command=root.destroy).pack(anchor="w", pady=(16, 0))
    root.mainloop()


def _open_child_tool(parent: tk.Tk, tool_id: str) -> None:
    try:
        launch_tool(tool_id, parent)
    except Exception as exc:
        messagebox.showerror("Error", str(exc))


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Neverbroken web archive toolkit")
    parser.add_argument(
        "--tool",
        choices=sorted(TOOL_DEFINITIONS.keys()),
        help="Open a specific workflow instead of the launcher.",
    )
    parser.add_argument(
        "--list-tools",
        action="store_true",
        help="List available workflow IDs and exit.",
    )
    args = parser.parse_args(argv)

    if args.list_tools:
        for tool_id, tool in TOOL_DEFINITIONS.items():
            print(f"{tool_id}: {tool['description']}")
        return 0

    ensure_project_directories()
    if args.tool:
        launch_tool(args.tool)
    else:
        launch_tool_selector()
    return 0

