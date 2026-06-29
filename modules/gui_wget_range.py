"""Tkinter workflow for downloading a row range from a snapshot workbook.

This GUI is the compatibility workflow for the older wget Excel script. It
lets the user choose an Excel file and row range, then delegates command
execution and log workbook creation to ``modules.downloads``.
"""

from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox

from modules.downloads import download_range_from_excel


class WgetRangeApp:
    """Desktop window for range-based wget downloads.

    Args:
        root: Tkinter root or child window that should contain this workflow.

    Returns:
        A configured ``WgetRangeApp`` instance bound to ``root``.
    """

    def __init__(self, root: tk.Tk | tk.Toplevel) -> None:
        """Build the range downloader interface.

        Args:
            root: Tkinter root or child window for this workflow.

        Returns:
            None. The constructor creates widgets and stores them on the
            instance.
        """
        # Store the window so callbacks can interact with the UI.
        self.root = root
        self.root.title("Broken Link Recovery Tool - Wget Range Downloader")

        # File picker row: the user selects the workbook containing capture URLs.
        tk.Label(root, text="Excel File:").grid(row=0, column=0, padx=10, pady=10)
        self.excel_file_entry = tk.Entry(root, width=50)
        self.excel_file_entry.grid(row=0, column=1, padx=10, pady=10)
        tk.Button(root, text="Browse", command=self.select_excel_file).grid(
            row=0,
            column=2,
            padx=10,
            pady=10,
        )

        # Start and end row inputs define an inclusive Excel row range.
        tk.Label(root, text="Start Row:").grid(row=1, column=0, padx=10, pady=10)
        self.start_row_entry = tk.Entry(root, width=10)
        self.start_row_entry.grid(row=1, column=1, padx=10, pady=10, sticky="w")

        tk.Label(root, text="End Row:").grid(row=2, column=0, padx=10, pady=10)
        self.end_row_entry = tk.Entry(root, width=10)
        self.end_row_entry.grid(row=2, column=1, padx=10, pady=10, sticky="w")

        # Button callback validates input and runs wget for each selected row.
        tk.Button(root, text="Download", command=self.download_range).grid(
            row=3,
            column=1,
            padx=10,
            pady=20,
        )

    def select_excel_file(self) -> None:
        """Open a file picker and store the selected workbook path.

        Args:
            None.

        Returns:
            None. The entry field is updated if a file is selected.
        """
        # Ask for an Excel workbook path that the range downloader can process.
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])

        # If the user chose a file, replace any previous entry value.
        if file_path:
            self.excel_file_entry.delete(0, tk.END)
            self.excel_file_entry.insert(0, file_path)

    def download_range(self) -> None:
        """Run wget for each capture link in the selected row range.

        Args:
            None. The function reads workbook and row values from UI fields.

        Returns:
            None. Success and error states are shown with Tkinter message boxes.
        """
        # Convert the entry text into a Path so validation can use filesystem
        # helpers.
        excel_file = Path(self.excel_file_entry.get())

        # If the selected path is missing or is a directory, stop before calling
        # the downloader.
        if not excel_file.is_file():
            messagebox.showerror("Error", "Please select a valid Excel file.")
            return

        try:
            # Convert the row inputs to integers because Excel rows are numeric.
            start_row = int(self.start_row_entry.get())
            end_row = int(self.end_row_entry.get())

            # Pass the validated workbook and row range into the shared download
            # utility.
            log_paths = download_range_from_excel(excel_file, start_row, end_row)
        except Exception as exc:
            # Show parsing, workbook, or wget errors to the user.
            messagebox.showerror("Error", str(exc))
            return

        # Report how many log workbooks were created for the selected range.
        messagebox.showinfo(
            "Success",
            f"Download and logging completed. Created {len(log_paths)} log file(s).",
        )


def launch(parent: tk.Tk | tk.Toplevel | None = None):
    """Launch the row-range wget workflow window.

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
    WgetRangeApp(window)

    # Only start a mainloop when this workflow owns the root window.
    if parent is None:
        window.mainloop()
    return window
