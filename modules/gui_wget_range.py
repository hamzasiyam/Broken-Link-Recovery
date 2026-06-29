from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox

from modules.downloads import download_range_from_excel


class WgetRangeApp:
    def __init__(self, root: tk.Tk | tk.Toplevel) -> None:
        self.root = root
        self.root.title("Wget Downloader")

        tk.Label(root, text="Excel File:").grid(row=0, column=0, padx=10, pady=10)
        self.excel_file_entry = tk.Entry(root, width=50)
        self.excel_file_entry.grid(row=0, column=1, padx=10, pady=10)
        tk.Button(root, text="Browse", command=self.select_excel_file).grid(
            row=0,
            column=2,
            padx=10,
            pady=10,
        )

        tk.Label(root, text="Start Row:").grid(row=1, column=0, padx=10, pady=10)
        self.start_row_entry = tk.Entry(root, width=10)
        self.start_row_entry.grid(row=1, column=1, padx=10, pady=10, sticky="w")

        tk.Label(root, text="End Row:").grid(row=2, column=0, padx=10, pady=10)
        self.end_row_entry = tk.Entry(root, width=10)
        self.end_row_entry.grid(row=2, column=1, padx=10, pady=10, sticky="w")

        tk.Button(root, text="Download", command=self.download_range).grid(
            row=3,
            column=1,
            padx=10,
            pady=20,
        )

    def select_excel_file(self) -> None:
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if file_path:
            self.excel_file_entry.delete(0, tk.END)
            self.excel_file_entry.insert(0, file_path)

    def download_range(self) -> None:
        excel_file = Path(self.excel_file_entry.get())
        if not excel_file.is_file():
            messagebox.showerror("Error", "Please select a valid Excel file.")
            return

        try:
            start_row = int(self.start_row_entry.get())
            end_row = int(self.end_row_entry.get())
            log_paths = download_range_from_excel(excel_file, start_row, end_row)
        except Exception as exc:
            messagebox.showerror("Error", str(exc))
            return

        messagebox.showinfo(
            "Success",
            f"Download and logging completed. Created {len(log_paths)} log file(s).",
        )


def launch(parent: tk.Tk | tk.Toplevel | None = None):
    window = tk.Toplevel(parent) if parent is not None else tk.Tk()
    WgetRangeApp(window)
    if parent is None:
        window.mainloop()
    return window
