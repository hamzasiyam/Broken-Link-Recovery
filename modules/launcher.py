"""Application launcher and command-line routing for Broken Link Recovery Tool.

This module owns the single entry point for the desktop application. It keeps
the list of available workflows in one place, opens the selector window, and
supports command-line arguments for launching a specific workflow directly.
"""

import argparse
import importlib
import tkinter as tk
from tkinter import messagebox

from modules.paths import ensure_project_directories


PROGRAM_NAME = "Broken Link Recovery Tool"

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
    """Launch one registered workflow window.

    Args:
        tool_id: A string key from ``TOOL_DEFINITIONS`` such as
            ``"snapshot-excel"`` or ``"archive-downloader"``.
        parent: Optional Tkinter root or parent window. If provided, the
            workflow opens as a child window instead of creating a new root.

    Returns:
        The Tkinter window object returned by the workflow module's ``launch``
        function.
    """
    # Look up the workflow metadata so the launcher does not need to import
    # every GUI module before it is actually used.
    tool = TOOL_DEFINITIONS[tool_id]

    # Import the selected GUI module dynamically, then call its standard
    # ``launch`` function with the optional parent window.
    module = importlib.import_module(tool["module"])
    return module.launch(parent)


def launch_tool_selector() -> None:
    """Open the main workflow selector window.

    Args:
        None.

    Returns:
        None. This function starts Tkinter's event loop and exits only after
        the user closes the selector window.
    """
    # Ensure the folders used by reports, profiles, and assets exist before
    # any workflow tries to read from or write to them.
    ensure_project_directories()

    root = tk.Tk()
    root.title(PROGRAM_NAME)
    root.geometry("460x310")

    # Create a simple selector frame so each workflow remains one click away.
    frame = tk.Frame(root, padx=16, pady=16)
    frame.pack(fill="both", expand=True)

    title = tk.Label(
        frame,
        text=PROGRAM_NAME,
        font=("Arial", 14, "bold"),
    )
    title.pack(anchor="w", pady=(0, 8))

    subtitle = tk.Label(
        frame,
        text="Choose a workflow to open.",
        font=("Arial", 10),
    )
    subtitle.pack(anchor="w", pady=(0, 12))

    # Build one button per registered workflow. The default argument on the
    # lambda captures the current loop value so every button opens its own tool.
    for tool_id, tool in TOOL_DEFINITIONS.items():
        button = tk.Button(
            frame,
            text=tool["label"],
            width=34,
            command=lambda selected_tool=tool_id: _open_child_tool(root, selected_tool),
        )
        button.pack(anchor="w", pady=4)

    # Close destroys only the selector root; child windows use Tkinter's normal
    # window controls.
    tk.Button(frame, text="Close", width=34, command=root.destroy).pack(anchor="w", pady=(16, 0))
    root.mainloop()


def _open_child_tool(parent: tk.Tk, tool_id: str) -> None:
    """Open a workflow from the selector and surface startup errors to users.

    Args:
        parent: The selector window that should own the child workflow.
        tool_id: A string key from ``TOOL_DEFINITIONS``.

    Returns:
        None. The function either opens a child window or shows an error dialog.
    """
    try:
        # Pass the selector root as the parent so the selected workflow opens as
        # a child window instead of starting a second root event loop.
        launch_tool(tool_id, parent)
    except Exception as exc:
        # If a workflow cannot start, show the exception in a GUI dialog instead
        # of letting Tkinter fail silently in the callback.
        messagebox.showerror("Error", str(exc))


def main(argv: list[str] | None = None) -> int:
    """Parse command-line options and start the requested workflow.

    Args:
        argv: Optional list of argument strings. Pass ``None`` to let argparse
            read from ``sys.argv`` when the program is run from the terminal.

    Returns:
        Integer process status code. ``0`` indicates successful argument
        handling and launcher startup.
    """
    # Configure the command-line interface in one place so both users and tests
    # can discover available workflows without opening a GUI.
    parser = argparse.ArgumentParser(description="Broken Link Recovery Tool")
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

    # If the user only wants a list of workflows, print the registry and exit
    # without opening any Tkinter windows.
    if args.list_tools:
        for tool_id, tool in TOOL_DEFINITIONS.items():
            print(f"{tool_id}: {tool['description']}")
        return 0

    # Create runtime folders before launching any workflow that may write files.
    ensure_project_directories()

    # If a specific workflow was requested, launch it directly; otherwise open
    # the selector so the user can choose interactively.
    if args.tool:
        launch_tool(args.tool)
    else:
        launch_tool_selector()
    return 0
