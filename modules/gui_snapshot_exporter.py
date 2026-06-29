import tkinter as tk
from tkinter import messagebox

from modules.excel_reports import save_snapshots_to_excel, snapshot_excel_path
from modules.wayback import domain_from_url, get_snapshots


class SnapshotExporterApp:
    def __init__(self, root: tk.Tk | tk.Toplevel) -> None:
        self.root = root
        self.root.title("Wayback Machine Snapshot Exporter")

        frame = tk.Frame(root, padx=10, pady=10)
        frame.pack(padx=10, pady=10)

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

        tk.Label(frame, text="Enter URL:").grid(row=1, column=0, pady=5)
        self.entry_url = tk.Entry(frame, width=50)
        self.entry_url.grid(row=1, column=1, pady=5)

        button_process = tk.Button(
            frame,
            text="Process Snapshots",
            command=self.process_snapshots,
        )
        button_process.grid(row=2, columnspan=2, pady=10)

    def process_snapshots(self) -> None:
        url = self.entry_url.get().strip()
        if not url:
            messagebox.showerror("Error", "Please enter a URL.")
            return

        try:
            snapshots = get_snapshots(url)
            if not snapshots:
                messagebox.showerror("Error", "No snapshots found.")
                return

            domain = domain_from_url(url)
            output_path = save_snapshots_to_excel(snapshots, snapshot_excel_path(domain))
        except Exception as exc:
            messagebox.showerror("Error", str(exc))
            return

        messagebox.showinfo("Success", f"Snapshots and summary saved to {output_path}")


def launch(parent: tk.Tk | tk.Toplevel | None = None):
    window = tk.Toplevel(parent) if parent is not None else tk.Tk()
    SnapshotExporterApp(window)
    if parent is None:
        window.mainloop()
    return window

