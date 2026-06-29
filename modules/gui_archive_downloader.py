import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import pandas as pd

from modules.downloads import WgetDownloader, check_wget, process_spreadsheet
from modules.profiles import proxy_profile_store
from modules.proxy import connect_proxy, current_proxy_info, disable_proxy


class ArchiveDownloaderApp:
    def __init__(self, root: tk.Tk | tk.Toplevel) -> None:
        self.root = root
        self.root.title("Web Archive Downloader with Proxy")
        self.downloader = WgetDownloader()

        self.proxy_type_var = tk.StringVar(value="HTTP")
        self.disable_proxy_var = tk.IntVar()
        self.download_delay_var = tk.StringVar(value="5")
        self.profile_name_var = tk.StringVar()

        self._build_ui()
        self.load_profile_names()
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)

    def _build_ui(self) -> None:
        frame = tk.Frame(self.root, padx=10, pady=10)
        frame.pack(padx=10, pady=10)

        summary_label = tk.Label(
            frame,
            text=(
                "This program downloads web content based on URLs from an Excel "
                "spreadsheet, supports proxy settings, and saves the downloaded "
                "content into organized directories."
            ),
            wraplength=400,
            justify="left",
            font=("Arial", 10, "italic"),
        )
        summary_label.grid(row=0, columnspan=3, pady=10)

        tk.Label(frame, text="Proxy Type:").grid(row=1, column=0, pady=5, sticky=tk.W)
        ttk.Combobox(
            frame,
            textvariable=self.proxy_type_var,
            values=["HTTP", "SOCKS5"],
            state="readonly",
        ).grid(row=1, column=1, pady=5)

        tk.Label(frame, text="Domain Name:").grid(row=2, column=0, pady=5, sticky=tk.W)
        self.entry_domain_name = tk.Entry(frame, width=50)
        self.entry_domain_name.grid(row=2, column=1, pady=5)

        tk.Label(frame, text="Proxy Port:").grid(row=3, column=0, pady=5, sticky=tk.W)
        self.entry_proxy_port = tk.Entry(frame, width=50)
        self.entry_proxy_port.grid(row=3, column=1, pady=5)

        tk.Label(frame, text="Proxy Username:").grid(row=4, column=0, pady=5, sticky=tk.W)
        self.entry_proxy_username = tk.Entry(frame, width=50)
        self.entry_proxy_username.grid(row=4, column=1, pady=5)

        tk.Label(frame, text="Proxy Password:").grid(row=5, column=0, pady=5, sticky=tk.W)
        self.entry_proxy_password = tk.Entry(frame, width=50, show="*")
        self.entry_proxy_password.grid(row=5, column=1, pady=5)

        tk.Checkbutton(
            frame,
            text="Disable Proxy",
            variable=self.disable_proxy_var,
        ).grid(row=6, columnspan=2, pady=5)

        tk.Label(frame, text="Time between downloads (seconds):").grid(
            row=7,
            column=0,
            pady=5,
            sticky=tk.W,
        )
        tk.Entry(frame, textvariable=self.download_delay_var, width=10).grid(
            row=7,
            column=1,
            pady=5,
            sticky=tk.W,
        )

        tk.Button(frame, text="Start", command=self.start_program).grid(
            row=8,
            columnspan=2,
            pady=10,
        )

        tk.Label(frame, text="Profile Name").grid(row=9, column=0, pady=10, sticky=tk.W)
        self.profile_name_combobox = ttk.Combobox(
            frame,
            textvariable=self.profile_name_var,
            state="readonly",
        )
        self.profile_name_combobox.grid(row=9, column=1, pady=10, padx=10)
        self.profile_name_combobox.bind("<<ComboboxSelected>>", self.select_profile_callback)

        tk.Button(frame, text="Create Profile", command=self.create_profile_window).grid(
            row=9,
            column=2,
            pady=10,
        )
        tk.Button(frame, text="Edit Profile", command=self.edit_profile_window).grid(
            row=10,
            column=1,
            pady=10,
        )
        tk.Button(frame, text="Delete Profile", command=self.delete_selected_profile).grid(
            row=10,
            column=2,
            pady=10,
        )

    def start_program(self) -> None:
        wget_available, wget_message = check_wget(self.downloader.wget_path)
        if not wget_available:
            messagebox.showerror("Error", f"wget is not available: {wget_message}")
            return

        if self.disable_proxy_var.get() == 1:
            disable_proxy()
        elif self.entry_domain_name.get().strip():
            connect_proxy(
                self.proxy_type_var.get(),
                self.entry_domain_name.get(),
                self.entry_proxy_port.get(),
                self.entry_proxy_username.get(),
                self.entry_proxy_password.get(),
            )
        else:
            disable_proxy()

        proxy_info = current_proxy_info()
        print(f"HTTP Proxy: {proxy_info['http_proxy']}")
        print(f"HTTPS Proxy: {proxy_info['https_proxy']}")
        print(wget_message)
        self.select_file()

    def select_file(self) -> None:
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if not file_path:
            return

        sheet_names = pd.ExcelFile(file_path).sheet_names
        self.sheet_selector_window(file_path, sheet_names)

    def sheet_selector_window(self, file_path: str, sheet_names: list[str]) -> None:
        sheet_window = tk.Toplevel(self.root)
        sheet_window.title("Select Sheet")

        sheet_name_var = tk.StringVar(value=sheet_names[0])
        tk.Label(sheet_window, text="Select Sheet:").pack(pady=10)
        ttk.Combobox(
            sheet_window,
            textvariable=sheet_name_var,
            values=sheet_names,
            state="readonly",
        ).pack(pady=10)

        def on_sheet_select() -> None:
            try:
                delay = int(self.download_delay_var.get())
                output_dir = os.path.join(os.path.dirname(file_path), "downloaded_contents")
                final_dir = process_spreadsheet(
                    file_path,
                    sheet_name_var.get(),
                    output_dir,
                    delay,
                    downloader=self.downloader,
                )
            except Exception as exc:
                messagebox.showerror("Error", f"Failed to process the spreadsheet: {exc}")
                return

            sheet_window.destroy()
            messagebox.showinfo("Success", f"Downloaded contents successfully to {final_dir}.")

        tk.Button(sheet_window, text="Select", command=on_sheet_select).pack(pady=20)

    def create_profile_window(self) -> None:
        self._profile_window("Create Profile")

    def edit_profile_window(self) -> None:
        selected_profile = self.profile_name_var.get()
        profiles = proxy_profile_store.load()
        if selected_profile not in profiles:
            messagebox.showerror("Error", "Selected profile does not exist.")
            return
        self._profile_window("Edit Profile", selected_profile, profiles[selected_profile])

    def _profile_window(
        self,
        title: str,
        selected_profile: str | None = None,
        profile_data: dict[str, str] | None = None,
    ) -> None:
        profile_window = tk.Toplevel(self.root)
        profile_window.title(title)

        profile_name = tk.StringVar(value=selected_profile or "")
        profile_proxy_type = tk.StringVar(value=(profile_data or {}).get("proxy_type", "HTTP"))
        profile_domain_name = tk.StringVar(value=(profile_data or {}).get("domain_name", ""))
        profile_proxy_port = tk.StringVar(value=(profile_data or {}).get("proxy_port", ""))
        profile_proxy_username = tk.StringVar(value=(profile_data or {}).get("proxy_username", ""))
        profile_proxy_password = tk.StringVar(value=(profile_data or {}).get("proxy_password", ""))

        tk.Label(profile_window, text="Profile Name:").grid(row=0, column=0, pady=10, sticky=tk.W)
        profile_name_state = "readonly" if selected_profile else "normal"
        tk.Entry(
            profile_window,
            textvariable=profile_name,
            width=50,
            state=profile_name_state,
        ).grid(row=0, column=1)

        tk.Label(profile_window, text="Proxy Type:").grid(row=1, column=0, pady=10, sticky=tk.W)
        ttk.Combobox(
            profile_window,
            textvariable=profile_proxy_type,
            values=["HTTP", "SOCKS5"],
            state="readonly",
        ).grid(row=1, column=1)

        tk.Label(profile_window, text="Domain Name:").grid(row=2, column=0, pady=10, sticky=tk.W)
        tk.Entry(profile_window, textvariable=profile_domain_name, width=50).grid(row=2, column=1)

        tk.Label(profile_window, text="Proxy Port:").grid(row=3, column=0, pady=10, sticky=tk.W)
        tk.Entry(profile_window, textvariable=profile_proxy_port, width=50).grid(row=3, column=1)

        tk.Label(profile_window, text="Proxy Username:").grid(row=4, column=0, pady=10, sticky=tk.W)
        tk.Entry(profile_window, textvariable=profile_proxy_username, width=50).grid(row=4, column=1)

        tk.Label(profile_window, text="Proxy Password:").grid(row=5, column=0, pady=10, sticky=tk.W)
        tk.Entry(
            profile_window,
            textvariable=profile_proxy_password,
            width=50,
            show="*",
        ).grid(row=5, column=1)

        def save_profile() -> None:
            profile = {
                "proxy_type": profile_proxy_type.get(),
                "domain_name": profile_domain_name.get(),
                "proxy_port": profile_proxy_port.get(),
                "proxy_username": profile_proxy_username.get(),
                "proxy_password": profile_proxy_password.get(),
            }
            try:
                proxy_profile_store.save(profile_name.get(), profile)
            except Exception as exc:
                messagebox.showerror("Error", str(exc))
                return
            profile_window.destroy()
            self.load_profile_names()

        tk.Button(profile_window, text="Save Profile", command=save_profile).grid(
            row=6,
            columnspan=2,
            pady=20,
        )

    def delete_selected_profile(self) -> None:
        selected_profile = self.profile_name_var.get()
        if not selected_profile:
            messagebox.showerror("Error", "Select a profile to delete.")
            return

        if messagebox.askyesno(
            "Delete Profile",
            f"Are you sure you want to delete the profile '{selected_profile}'?",
        ):
            proxy_profile_store.delete(selected_profile)
            self.load_profile_names()
            self.clear_profile_fields()

    def clear_profile_fields(self) -> None:
        self.proxy_type_var.set("HTTP")
        for entry in (
            self.entry_domain_name,
            self.entry_proxy_port,
            self.entry_proxy_username,
            self.entry_proxy_password,
        ):
            entry.delete(0, tk.END)

    def load_profile_names(self) -> None:
        self.profile_name_combobox["values"] = proxy_profile_store.names()

    def select_profile_callback(self, _event) -> None:
        profiles = proxy_profile_store.load()
        profile_name = self.profile_name_var.get()
        if profile_name not in profiles:
            return

        profile = profiles[profile_name]
        self.proxy_type_var.set(profile["proxy_type"])
        self.entry_domain_name.delete(0, tk.END)
        self.entry_domain_name.insert(0, profile["domain_name"])
        self.entry_proxy_port.delete(0, tk.END)
        self.entry_proxy_port.insert(0, profile["proxy_port"])
        self.entry_proxy_username.delete(0, tk.END)
        self.entry_proxy_username.insert(0, profile["proxy_username"])
        self.entry_proxy_password.delete(0, tk.END)
        self.entry_proxy_password.insert(0, profile["proxy_password"])

    def on_closing(self) -> None:
        self.downloader.terminate()
        self.root.destroy()


def launch(parent: tk.Tk | tk.Toplevel | None = None):
    window = tk.Toplevel(parent) if parent is not None else tk.Tk()
    ArchiveDownloaderApp(window)
    if parent is None:
        window.mainloop()
    return window

