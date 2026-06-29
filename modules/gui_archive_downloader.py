"""Tkinter workflow for downloading archived website files with optional proxy.

This module owns the archive-downloader user interface: proxy settings,
download delay, profile management, workbook selection, and sheet selection.
Actual wget execution and spreadsheet processing live in shared modules.
"""

import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import pandas as pd

from modules.downloads import WgetDownloader, check_wget, process_spreadsheet
from modules.profiles import proxy_profile_store
from modules.proxy import connect_proxy, current_proxy_info, disable_proxy


class ArchiveDownloaderApp:
    """Desktop window for downloading archived files from snapshot workbooks.

    Args:
        root: Tkinter root or child window that should contain this workflow.

    Returns:
        A configured ``ArchiveDownloaderApp`` instance bound to ``root``.
    """

    def __init__(self, root: tk.Tk | tk.Toplevel) -> None:
        """Initialize downloader state and build the archive-download UI.

        Args:
            root: Tkinter root or child window for this workflow.

        Returns:
            None. The constructor creates widgets, profile state, and close
            handling.
        """
        # Store the root window so callbacks can create dialogs and child
        # windows.
        self.root = root
        self.root.title("Broken Link Recovery Tool - Archive Downloader")

        # Keep one downloader instance so the active wget process can be
        # terminated if the user closes the window.
        self.downloader = WgetDownloader()

        # Tkinter variables hold form state used by callbacks.
        self.proxy_type_var = tk.StringVar(value="HTTP")
        self.disable_proxy_var = tk.IntVar()
        self.download_delay_var = tk.StringVar(value="5")
        self.profile_name_var = tk.StringVar()

        # Build the interface, load saved proxy profiles, and wire close cleanup.
        self._build_ui()
        self.load_profile_names()
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)

    def _build_ui(self) -> None:
        """Create all widgets for the archive downloader window.

        Args:
            None.

        Returns:
            None. Widgets are added to ``self.root``.
        """
        # Use one padded frame to keep related proxy/download controls together.
        frame = tk.Frame(self.root, padx=10, pady=10)
        frame.pack(padx=10, pady=10)

        # Explain the workflow before the user starts configuring proxy values.
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

        # Proxy controls feed environment variables consumed by wget.
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

        # Disable Proxy lets users explicitly clear proxy environment variables.
        tk.Checkbutton(
            frame,
            text="Disable Proxy",
            variable=self.disable_proxy_var,
        ).grid(row=6, columnspan=2, pady=5)

        # Delay helps avoid aggressive back-to-back downloads.
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

        # Start validates wget/proxy settings, then opens the workbook selector.
        tk.Button(frame, text="Start", command=self.start_program).grid(
            row=8,
            columnspan=2,
            pady=10,
        )

        # Proxy profile selector and management buttons.
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
        """Validate wget, apply proxy settings, and start workbook selection.

        Args:
            None. Values are read from the proxy and delay widgets.

        Returns:
            None. The function either opens the file selector or shows an error.
        """
        # Check wget before asking for a workbook so users get dependency errors
        # immediately.
        wget_available, wget_message = check_wget(self.downloader.wget_path)

        # If wget cannot run, stop before any network or file processing begins.
        if not wget_available:
            messagebox.showerror("Error", f"wget is not available: {wget_message}")
            return

        # If the disable checkbox is selected, clear proxy environment variables.
        if self.disable_proxy_var.get() == 1:
            disable_proxy()

        # If proxy fields contain a domain, configure the proxy for wget.
        elif self.entry_domain_name.get().strip():
            connect_proxy(
                self.proxy_type_var.get(),
                self.entry_domain_name.get(),
                self.entry_proxy_port.get(),
                self.entry_proxy_username.get(),
                self.entry_proxy_password.get(),
            )

        # If no proxy is configured and disable is not checked, still clear
        # existing proxy variables so stale values do not affect downloads.
        else:
            disable_proxy()

        # Print proxy and wget details to the terminal for troubleshooting.
        proxy_info = current_proxy_info()
        print(f"HTTP Proxy: {proxy_info['http_proxy']}")
        print(f"HTTPS Proxy: {proxy_info['https_proxy']}")
        print(wget_message)

        # Continue to workbook selection after environment setup succeeds.
        self.select_file()

    def select_file(self) -> None:
        """Prompt the user for an Excel workbook and open sheet selection.

        Args:
            None.

        Returns:
            None. The sheet selector opens only when a file is selected.
        """
        # Ask for the workbook produced by the snapshot Excel workflow.
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])

        # If the user cancels the file dialog, stop without showing an error.
        if not file_path:
            return

        # Read available sheet names so the user can choose the correct sheet.
        sheet_names = pd.ExcelFile(file_path).sheet_names
        self.sheet_selector_window(file_path, sheet_names)

    def sheet_selector_window(self, file_path: str, sheet_names: list[str]) -> None:
        """Open a child dialog for choosing which workbook sheet to process.

        Args:
            file_path: Path string to the selected Excel workbook.
            sheet_names: List of worksheet name strings from the workbook.

        Returns:
            None. The selected sheet is processed from a nested callback.
        """
        # Create a child window so sheet selection stays tied to this workflow.
        sheet_window = tk.Toplevel(self.root)
        sheet_window.title("Select Sheet")

        # Default to the first sheet while still allowing the user to choose.
        sheet_name_var = tk.StringVar(value=sheet_names[0])
        tk.Label(sheet_window, text="Select Sheet:").pack(pady=10)
        ttk.Combobox(
            sheet_window,
            textvariable=sheet_name_var,
            values=sheet_names,
            state="readonly",
        ).pack(pady=10)

        def on_sheet_select() -> None:
            """Process the selected worksheet and download its captures.

            Args:
                None. The function reads selected sheet and delay values from
                enclosing Tkinter variables.

            Returns:
                None. The dialog closes on success or shows an error message.
            """
            try:
                # Convert the delay field to an integer before passing it to the
                # spreadsheet processor.
                delay = int(self.download_delay_var.get())

                # Store downloads beside the chosen workbook so outputs are easy
                # to find for that spreadsheet.
                output_dir = os.path.join(os.path.dirname(file_path), "downloaded_contents")

                # Delegate row processing and wget execution to the shared
                # downloads module.
                final_dir = process_spreadsheet(
                    file_path,
                    sheet_name_var.get(),
                    output_dir,
                    delay,
                    downloader=self.downloader,
                )
            except Exception as exc:
                # If sheet parsing, delay parsing, or downloading fails, keep the
                # sheet selector open and show the reason.
                messagebox.showerror("Error", f"Failed to process the spreadsheet: {exc}")
                return

            # Close sheet selection and show the final download directory.
            sheet_window.destroy()
            messagebox.showinfo("Success", f"Downloaded contents successfully to {final_dir}.")

        # Select button runs the nested callback above.
        tk.Button(sheet_window, text="Select", command=on_sheet_select).pack(pady=20)

    def create_profile_window(self) -> None:
        """Open the proxy profile creation dialog.

        Args:
            None.

        Returns:
            None. The shared profile-window builder handles the dialog.
        """
        # Pass only a title because a new profile has no existing data.
        self._profile_window("Create Profile")

    def edit_profile_window(self) -> None:
        """Open the proxy profile edit dialog for the selected profile.

        Args:
            None. The selected profile name is read from the combobox variable.

        Returns:
            None. An error dialog is shown if no matching profile exists.
        """
        # Load saved profiles so the selected one can be validated and edited.
        selected_profile = self.profile_name_var.get()
        profiles = proxy_profile_store.load()

        # If the selected profile is absent, stop before opening the editor.
        if selected_profile not in profiles:
            messagebox.showerror("Error", "Selected profile does not exist.")
            return

        # Open the editor populated with existing profile values.
        self._profile_window("Edit Profile", selected_profile, profiles[selected_profile])

    def _profile_window(
        self,
        title: str,
        selected_profile: str | None = None,
        profile_data: dict[str, str] | None = None,
    ) -> None:
        """Build the create/edit dialog for proxy profiles.

        Args:
            title: Dialog title string.
            selected_profile: Optional existing profile name string. When
                provided, the profile name field is read-only.
            profile_data: Optional dictionary of existing proxy settings.

        Returns:
            None. Profile data is saved from the nested callback.
        """
        # Create a child dialog for proxy profile fields.
        profile_window = tk.Toplevel(self.root)
        profile_window.title(title)

        # Populate form variables from existing data or sensible defaults.
        profile_name = tk.StringVar(value=selected_profile or "")
        profile_proxy_type = tk.StringVar(value=(profile_data or {}).get("proxy_type", "HTTP"))
        profile_domain_name = tk.StringVar(value=(profile_data or {}).get("domain_name", ""))
        profile_proxy_port = tk.StringVar(value=(profile_data or {}).get("proxy_port", ""))
        profile_proxy_username = tk.StringVar(value=(profile_data or {}).get("proxy_username", ""))
        profile_proxy_password = tk.StringVar(value=(profile_data or {}).get("proxy_password", ""))

        # Build the proxy profile editor fields.
        tk.Label(profile_window, text="Profile Name:").grid(row=0, column=0, pady=10, sticky=tk.W)

        # Existing profile names are read-only so editing cannot accidentally
        # duplicate a profile under a new key.
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
            """Persist the proxy profile editor fields to JSON.

            Args:
                None. The function reads local Tkinter variables from the
                enclosing dialog.

            Returns:
                None. The profile JSON file is updated, or an error dialog is
                shown.
            """
            # Collect proxy form values into the JSON structure expected by the
            # profile store.
            profile = {
                "proxy_type": profile_proxy_type.get(),
                "domain_name": profile_domain_name.get(),
                "proxy_port": profile_proxy_port.get(),
                "proxy_username": profile_proxy_username.get(),
                "proxy_password": profile_proxy_password.get(),
            }
            try:
                # Save the profile by name, creating or replacing the JSON entry.
                proxy_profile_store.save(profile_name.get(), profile)
            except Exception as exc:
                # If saving fails, keep the dialog open and show the error.
                messagebox.showerror("Error", str(exc))
                return

            # Close the editor and refresh profile choices in the main window.
            profile_window.destroy()
            self.load_profile_names()

        # Save button runs the nested callback above.
        tk.Button(profile_window, text="Save Profile", command=save_profile).grid(
            row=6,
            columnspan=2,
            pady=20,
        )

    def delete_selected_profile(self) -> None:
        """Delete the selected proxy profile after user confirmation.

        Args:
            None. The selected profile name is read from the combobox variable.

        Returns:
            None. Profile fields and combobox values are updated on success.
        """
        # Read the selected profile name from the combobox.
        selected_profile = self.profile_name_var.get()

        # If no profile is selected, stop and ask the user to choose one.
        if not selected_profile:
            messagebox.showerror("Error", "Select a profile to delete.")
            return

        # Ask for confirmation because deleting a proxy profile cannot be undone
        # inside the app.
        if messagebox.askyesno(
            "Delete Profile",
            f"Are you sure you want to delete the profile '{selected_profile}'?",
        ):
            # Delete the profile, refresh combobox values, and clear the form.
            proxy_profile_store.delete(selected_profile)
            self.load_profile_names()
            self.clear_profile_fields()

    def clear_profile_fields(self) -> None:
        """Reset proxy form fields to their default empty state.

        Args:
            None.

        Returns:
            None. Tkinter variables and entry widgets are modified in place.
        """
        # Reset proxy type to the default selection.
        self.proxy_type_var.set("HTTP")

        # Clear every text entry used by proxy settings.
        for entry in (
            self.entry_domain_name,
            self.entry_proxy_port,
            self.entry_proxy_username,
            self.entry_proxy_password,
        ):
            entry.delete(0, tk.END)

    def load_profile_names(self) -> None:
        """Refresh the proxy-profile combobox with saved profile names.

        Args:
            None.

        Returns:
            None. The combobox values are replaced in place.
        """
        # Load profile names from JSON and pass them directly to the combobox.
        self.profile_name_combobox["values"] = proxy_profile_store.names()

    def select_profile_callback(self, _event) -> None:
        """Load the selected proxy profile into the proxy form fields.

        Args:
            _event: Tkinter combobox event object. It is unused because the
                selected profile name is read from ``self.profile_name_var``.

        Returns:
            None. Tkinter variables and entry widgets are updated in place.
        """
        # Load current profile data so selection reflects the latest JSON file.
        profiles = proxy_profile_store.load()
        profile_name = self.profile_name_var.get()

        # If the selected name is absent, ignore the event.
        if profile_name not in profiles:
            return

        # Copy stored profile values into the proxy form.
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
        """Terminate active downloads and close the window.

        Args:
            None.

        Returns:
            None. Any active wget process is terminated before destroying the
            root window.
        """
        # Stop any active wget process so closing the GUI does not leave a
        # background download running.
        self.downloader.terminate()

        # Destroy the Tkinter root or child window after cleanup.
        self.root.destroy()


def launch(parent: tk.Tk | tk.Toplevel | None = None):
    """Launch the archive downloader workflow window.

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
    ArchiveDownloaderApp(window)

    # Only start a mainloop when this workflow owns the root window.
    if parent is None:
        window.mainloop()
    return window
