"""Tkinter workflow for exporting Wayback snapshots to Excel.

This module owns only the user interface for the snapshot spreadsheet workflow.
It collects a URL from the user, then delegates snapshot fetching and workbook
creation to shared modules.
"""

import tkinter as tk
from tkinter import messagebox

from modules.excel_reports import save_snapshots_to_excel, snapshot_excel_path
from modules.wayback import domain_from_url, get_snapshots


class SnapshotExporterApp:
    """Desktop window for creating a snapshot Excel workbook.

    Args:
        root: Tkinter root or child window that should contain this workflow.

    Returns:
        A configured ``SnapshotExporterApp`` instance bound to ``root``.
    """

    def __init__(self, root: tk.Tk | tk.Toplevel) -> None:
        """Build the snapshot exporter interface.

        Args:
            root: Tkinter root or child window for this workflow.

        Returns:
            None. The constructor creates widgets and stores them on the
            instance.
        """
        # Store the root so callbacks can read widget values and show dialogs.
        self.root = root
        self.root.title("Broken Link Recovery Tool - Snapshot Excel Exporter")

        # Use one padded frame to keep the window layout compact and readable.
        frame = tk.Frame(root, padx=10, pady=10)
        frame.pack(padx=10, pady=10)

        # Explain the workflow directly in the window before the user enters a
        # URL.
        summary_label = tk.Label(
            frame,
            text=(
                "This tool retrieves historical snapshots of a website from the "
                "Wayback Machine and exports them into an Excel file with date "
                "formatting, HTTP links, and a summary."
            ),
            wraplength=400,
            justify="left",
            font=("Arial", 10, "italic"),
        )
        summary_label.grid(row=0, columnspan=2, pady=10)

        # URL entry field that process_snapshots reads when the button is
        # clicked.
        tk.Label(frame, text="Enter URL:").grid(row=1, column=0, pady=5)
        self.entry_url = tk.Entry(frame, width=50)
        self.entry_url.grid(row=1, column=1, pady=5)

        # Button callback performs validation, snapshot lookup, and Excel export.
        button_process = tk.Button(
            frame,
            text="Process Snapshots",
            command=self.process_snapshots,
        )
        button_process.grid(row=2, columnspan=2, pady=10)

    def process_snapshots(self) -> None:
        """Validate the URL, fetch snapshots, and save the Excel workbook.

        Args:
            None. The function reads the URL from ``self.entry_url``.

        Returns:
            None. Success and error states are shown with Tkinter message boxes.
        """
        # Read and trim the user-entered URL before validating it.
        url = self.entry_url.get().strip()

        # If the URL field is empty, stop and ask the user to provide one.
        if not url:
            messagebox.showerror("Error", "Please enter a URL.")
            return

        try:
            # Pass the URL into the shared Wayback helper to collect capture
            # links.
            snapshots = get_snapshots(url)

            # If waybackpack returned no captures, tell the user no workbook was
            # created.
            if not snapshots:
                messagebox.showerror("Error", "No snapshots found.")
                return

            # Derive the domain for the output filename, then pass snapshots to
            # the Excel writer.
            domain = domain_from_url(url)
            output_path = save_snapshots_to_excel(snapshots, snapshot_excel_path(domain))
        except Exception as exc:
            # Convert lower-level exceptions into a friendly GUI error dialog.
            messagebox.showerror("Error", str(exc))
            return

        # If every step completed, show the saved workbook path.
        messagebox.showinfo("Success", f"Snapshots and summary saved to {output_path}")


def launch(parent: tk.Tk | tk.Toplevel | None = None):
    """Launch the snapshot exporter workflow window.

    Args:
        parent: Optional Tkinter root or parent window. If provided, a child
            ``Toplevel`` is created. If omitted, this function creates a root.

    Returns:
        The Tkinter window object that hosts the workflow.
    """
    # If a parent exists, create a child window; otherwise create a standalone
    # root window for legacy script usage.
    window = tk.Toplevel(parent) if parent is not None else tk.Tk()

    # Instantiate the workflow class to populate the window with widgets.
    SnapshotExporterApp(window)

    # Only start a mainloop when this workflow owns the root window.
    if parent is None:
        window.mainloop()
    return window
