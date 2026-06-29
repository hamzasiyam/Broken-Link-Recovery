import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from modules.document_reports import (
    create_detailed_analysis_report,
    create_summary_report,
)
from modules.paths import LOGOS_DIR, PROCESSED_DIR
from modules.profiles import snapshot_profile_store
from modules.wayback import domain_from_url, get_snapshots, normalize_url


class SnapshotReportApp:
    def __init__(self, root: tk.Tk | tk.Toplevel) -> None:
        self.root = root
        self.root.title("Snapshot Analysis Report Generator")
        self.root.geometry("600x600")

        self.url_var = tk.StringVar()
        self.logo_path_var = tk.StringVar()
        self.logo_height_percent_var = tk.StringVar(value="50")
        self.title_color_hex_var = tk.StringVar(value="000000")
        self.heading_color_hex_var = tk.StringVar(value="000000")
        self.column_color_hex_var = tk.StringVar(value="FFFFFF")
        self.profile_name_var = tk.StringVar()

        self._build_ui()
        self.load_profile_names()

    def _build_ui(self) -> None:
        summary_label = tk.Label(
            self.root,
            text=(
                "This tool generates snapshot analysis reports for websites. It retrieves "
                "historical snapshots from a URL and creates detailed reports that help "
                "identify and fix broken links."
            ),
            wraplength=500,
            justify="left",
            font=("Arial", 10, "italic"),
        )
        summary_label.grid(row=0, columnspan=3, padx=10, pady=10)

        tk.Label(self.root, text="Enter URL:").grid(
            row=1,
            column=0,
            padx=10,
            pady=10,
            sticky="w",
        )
        tk.Entry(self.root, textvariable=self.url_var, width=50).grid(
            row=1,
            column=1,
            padx=10,
            pady=10,
            sticky="w",
        )

        tk.Label(self.root, text="Logo File:").grid(
            row=2,
            column=0,
            padx=10,
            pady=10,
            sticky="w",
        )
        tk.Entry(self.root, textvariable=self.logo_path_var, width=50).grid(
            row=2,
            column=1,
            padx=10,
            pady=10,
            sticky="w",
        )
        tk.Button(
            self.root,
            text="Browse",
            command=lambda: self.browse_file(self.logo_path_var),
        ).grid(row=2, column=2, padx=10, pady=10, sticky="w")

        tk.Label(self.root, text="Logo Height (%):").grid(
            row=3,
            column=0,
            padx=10,
            pady=10,
            sticky="w",
        )
        tk.Entry(self.root, textvariable=self.logo_height_percent_var, width=10).grid(
            row=3,
            column=1,
            padx=10,
            pady=10,
            sticky="w",
        )

        tk.Label(self.root, text="Title Font Color (Hex):").grid(
            row=4,
            column=0,
            padx=10,
            pady=10,
            sticky="w",
        )
        tk.Entry(self.root, textvariable=self.title_color_hex_var, width=10).grid(
            row=4,
            column=1,
            padx=10,
            pady=10,
            sticky="w",
        )

        tk.Label(self.root, text="Heading Font Color (Hex):").grid(
            row=5,
            column=0,
            padx=10,
            pady=10,
            sticky="w",
        )
        tk.Entry(self.root, textvariable=self.heading_color_hex_var, width=10).grid(
            row=5,
            column=1,
            padx=10,
            pady=10,
            sticky="w",
        )

        tk.Label(self.root, text="Column Shading Color (Hex):").grid(
            row=6,
            column=0,
            padx=10,
            pady=10,
            sticky="w",
        )
        tk.Entry(self.root, textvariable=self.column_color_hex_var, width=10).grid(
            row=6,
            column=1,
            padx=10,
            pady=10,
            sticky="w",
        )

        tk.Button(
            self.root,
            text="Generate Reports",
            command=self.generate_reports,
        ).grid(row=7, columnspan=3, padx=10, pady=20)

        tk.Label(self.root, text="Profile Name").grid(row=8, column=0, pady=10, sticky="w")
        self.profile_name_combobox = ttk.Combobox(
            self.root,
            textvariable=self.profile_name_var,
            state="readonly",
        )
        self.profile_name_combobox.grid(row=8, column=1, pady=10, padx=10)
        self.profile_name_combobox.bind("<<ComboboxSelected>>", self.select_profile_callback)

        tk.Button(self.root, text="Create Profile", command=self.create_profile_window).grid(
            row=9,
            column=0,
            pady=10,
        )
        tk.Button(
            self.root,
            text="Edit Profile",
            command=lambda: self.create_profile_window(existing_profile=self.profile_name_var.get()),
        ).grid(row=9, column=1, pady=10)
        tk.Button(self.root, text="Delete Profile", command=self.delete_profile_callback).grid(
            row=9,
            column=2,
            pady=10,
        )

    def browse_file(self, variable: tk.StringVar) -> None:
        file_path = filedialog.askopenfilename(
            initialdir=str(LOGOS_DIR),
            filetypes=[("Image files", "*.png *.jpg *.jpeg *.bmp *.gif")],
        )
        if file_path:
            variable.set(file_path)

    def generate_reports(self) -> None:
        try:
            url = normalize_url(self.url_var.get())
            domain = domain_from_url(url)
            snapshots = get_snapshots(url)
            if not snapshots:
                messagebox.showerror("Error", "No snapshots found.")
                return

            create_summary_report(
                snapshots,
                domain,
                self.logo_path_var.get(),
                self.logo_height_percent_var.get(),
                self.title_color_hex_var.get(),
                self.heading_color_hex_var.get(),
            )
            create_detailed_analysis_report(
                snapshots,
                domain,
                self.logo_path_var.get(),
                self.logo_height_percent_var.get(),
                self.title_color_hex_var.get(),
                self.heading_color_hex_var.get(),
                self.column_color_hex_var.get(),
            )
        except Exception as exc:
            messagebox.showerror("Error", str(exc))
            return

        messagebox.showinfo("Success", f"Reports saved to {PROCESSED_DIR}")

    def create_profile_window(self, existing_profile: str | None = None) -> None:
        if existing_profile == "":
            existing_profile = None

        profile_window = tk.Toplevel(self.root)
        profile_window.title("Create/Edit Profile")
        profile_window.geometry("550x450")

        profile_name = tk.StringVar(value=existing_profile or "")
        profile_url = tk.StringVar()
        profile_logo_path = tk.StringVar()
        profile_logo_height_percent = tk.StringVar(value="50")
        profile_title_color_hex = tk.StringVar(value="000000")
        profile_heading_color_hex = tk.StringVar(value="000000")
        profile_column_color_hex = tk.StringVar(value="FFFFFF")

        if existing_profile:
            profiles = snapshot_profile_store.load()
            if existing_profile not in profiles:
                messagebox.showerror("Error", "Selected profile does not exist.")
                profile_window.destroy()
                return
            profile = profiles[existing_profile]
            profile_url.set(profile["url"])
            profile_logo_path.set(profile["logo_path"])
            profile_logo_height_percent.set(profile["logo_height_percent"])
            profile_title_color_hex.set(profile["title_color_hex"])
            profile_heading_color_hex.set(profile["heading_color_hex"])
            profile_column_color_hex.set(profile["column_color_hex"])

        tk.Label(profile_window, text="Profile Name").grid(row=0, column=0, pady=10, sticky="w")
        tk.Entry(profile_window, textvariable=profile_name, width=50).grid(row=0, column=1)

        tk.Label(profile_window, text="URL").grid(row=1, column=0, pady=10, sticky="w")
        tk.Entry(profile_window, textvariable=profile_url, width=50).grid(row=1, column=1)

        tk.Label(profile_window, text="Logo File").grid(row=2, column=0, pady=10, sticky="w")
        tk.Entry(profile_window, textvariable=profile_logo_path, width=50).grid(row=2, column=1)
        tk.Button(
            profile_window,
            text="Browse",
            command=lambda: self.browse_file(profile_logo_path),
        ).grid(row=2, column=2)

        tk.Label(profile_window, text="Logo Height (%)").grid(
            row=3,
            column=0,
            pady=10,
            sticky="w",
        )
        tk.Entry(profile_window, textvariable=profile_logo_height_percent, width=10).grid(
            row=3,
            column=1,
            sticky="w",
        )

        tk.Label(profile_window, text="Title Color (hex)").grid(
            row=4,
            column=0,
            pady=10,
            sticky="w",
        )
        tk.Entry(profile_window, textvariable=profile_title_color_hex, width=10).grid(
            row=4,
            column=1,
            sticky="w",
        )

        tk.Label(profile_window, text="Heading Color (hex)").grid(
            row=5,
            column=0,
            pady=10,
            sticky="w",
        )
        tk.Entry(profile_window, textvariable=profile_heading_color_hex, width=10).grid(
            row=5,
            column=1,
            sticky="w",
        )

        tk.Label(profile_window, text="Column Color (hex)").grid(
            row=6,
            column=0,
            pady=10,
            sticky="w",
        )
        tk.Entry(profile_window, textvariable=profile_column_color_hex, width=10).grid(
            row=6,
            column=1,
            sticky="w",
        )

        def save_new_profile() -> None:
            profile_data = {
                "url": profile_url.get(),
                "logo_path": profile_logo_path.get(),
                "logo_height_percent": profile_logo_height_percent.get(),
                "title_color_hex": profile_title_color_hex.get(),
                "heading_color_hex": profile_heading_color_hex.get(),
                "column_color_hex": profile_column_color_hex.get(),
            }
            try:
                snapshot_profile_store.save(profile_name.get(), profile_data)
            except Exception as exc:
                messagebox.showerror("Error", str(exc))
                return
            profile_window.destroy()
            self.load_profile_names()

        tk.Button(profile_window, text="Save Profile", command=save_new_profile).grid(
            row=7,
            columnspan=3,
            pady=20,
        )

    def load_profile_names(self) -> None:
        self.profile_name_combobox["values"] = snapshot_profile_store.names()

    def select_profile_callback(self, _event) -> None:
        profiles = snapshot_profile_store.load()
        profile_name = self.profile_name_var.get()
        if profile_name not in profiles:
            return

        profile = profiles[profile_name]
        self.url_var.set(profile["url"])
        self.logo_path_var.set(profile["logo_path"])
        self.logo_height_percent_var.set(profile["logo_height_percent"])
        self.title_color_hex_var.set(profile["title_color_hex"])
        self.heading_color_hex_var.set(profile["heading_color_hex"])
        self.column_color_hex_var.set(profile["column_color_hex"])

    def delete_profile_callback(self) -> None:
        if snapshot_profile_store.delete(self.profile_name_var.get()):
            self.load_profile_names()
            messagebox.showinfo("Success", "Profile deleted successfully.")
        else:
            messagebox.showerror("Error", "Profile not found.")


def launch(parent: tk.Tk | tk.Toplevel | None = None):
    window = tk.Toplevel(parent) if parent is not None else tk.Tk()
    SnapshotReportApp(window)
    if parent is None:
        window.mainloop()
    return window
