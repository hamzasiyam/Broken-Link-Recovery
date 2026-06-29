"""Tkinter workflow for generating Word snapshot reports.

This module owns the report-generator user interface: URL input, optional logo
selection, report colors, and saved style profiles. It delegates actual
Wayback lookup and Word document creation to shared modules.
"""

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
    """Desktop window for creating snapshot analysis Word reports.

    Args:
        root: Tkinter root or child window that should contain this workflow.

    Returns:
        A configured ``SnapshotReportApp`` instance bound to ``root``.
    """

    def __init__(self, root: tk.Tk | tk.Toplevel) -> None:
        """Initialize report-generator state and build the interface.

        Args:
            root: Tkinter root or child window for this workflow.

        Returns:
            None. The constructor creates Tkinter variables and widgets.
        """
        # Store the root window so callbacks can create dialogs and child
        # windows.
        self.root = root
        self.root.title("Broken Link Recovery Tool - Snapshot Report Generator")
        self.root.geometry("600x600")

        # Tkinter variables keep widget state available to callback methods.
        self.url_var = tk.StringVar()
        self.logo_path_var = tk.StringVar()
        self.logo_height_percent_var = tk.StringVar(value="50")
        self.title_color_hex_var = tk.StringVar(value="000000")
        self.heading_color_hex_var = tk.StringVar(value="000000")
        self.column_color_hex_var = tk.StringVar(value="FFFFFF")
        self.profile_name_var = tk.StringVar()

        # Build widgets first, then load saved profile names into the combobox.
        self._build_ui()
        self.load_profile_names()

    def _build_ui(self) -> None:
        """Create all widgets for the report-generator window.

        Args:
            None.

        Returns:
            None. Widgets are added to ``self.root``.
        """
        # Summary text explains the workflow before the user enters settings.
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

        # URL field controls which website will be queried through Wayback.
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

        # Logo path field optionally brands the generated Word documents.
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

        # Logo height percentage lets users scale an image without editing the
        # image file itself.
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

        # Hex color fields feed directly into python-docx color formatting.
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

        # The generate button is the main action for creating report files.
        tk.Button(
            self.root,
            text="Generate Reports",
            command=self.generate_reports,
        ).grid(row=7, columnspan=3, padx=10, pady=20)

        # Profile selector lets users reuse saved report settings.
        tk.Label(self.root, text="Profile Name").grid(row=8, column=0, pady=10, sticky="w")
        self.profile_name_combobox = ttk.Combobox(
            self.root,
            textvariable=self.profile_name_var,
            state="readonly",
        )
        self.profile_name_combobox.grid(row=8, column=1, pady=10, padx=10)
        self.profile_name_combobox.bind("<<ComboboxSelected>>", self.select_profile_callback)

        # Profile buttons open the editor, edit the selected entry, or delete it.
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
        """Open an image picker and store the selected path in a Tk variable.

        Args:
            variable: ``tk.StringVar`` that should receive the selected image
                path.

        Returns:
            None. The variable is updated only when the user selects a file.
        """
        # Start in the known logos folder and restrict choices to common image
        # formats supported by Pillow.
        file_path = filedialog.askopenfilename(
            initialdir=str(LOGOS_DIR),
            filetypes=[("Image files", "*.png *.jpg *.jpeg *.bmp *.gif")],
        )

        # If the user selected an image, store it for report generation.
        if file_path:
            variable.set(file_path)

    def generate_reports(self) -> None:
        """Fetch snapshots and create summary and detailed Word reports.

        Args:
            None. The function reads report settings from Tkinter variables.

        Returns:
            None. Success and error states are shown with Tkinter message boxes.
        """
        try:
            # Normalize the URL before deriving the domain or calling Wayback.
            url = normalize_url(self.url_var.get())
            domain = domain_from_url(url)

            # Pass the URL to the shared Wayback helper to fetch capture links.
            snapshots = get_snapshots(url)

            # If no snapshots were found, stop before creating empty documents.
            if not snapshots:
                messagebox.showerror("Error", "No snapshots found.")
                return

            # Create the summary report using the current UI settings.
            create_summary_report(
                snapshots,
                domain,
                self.logo_path_var.get(),
                self.logo_height_percent_var.get(),
                self.title_color_hex_var.get(),
                self.heading_color_hex_var.get(),
            )

            # Create the detailed report using the same settings plus table color.
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
            # Show validation, Wayback, image, or document errors in the GUI.
            messagebox.showerror("Error", str(exc))
            return

        # If both documents were created, report the output directory.
        messagebox.showinfo("Success", f"Reports saved to {PROCESSED_DIR}")

    def create_profile_window(self, existing_profile: str | None = None) -> None:
        """Open the create/edit dialog for report-style profiles.

        Args:
            existing_profile: Optional profile name string. If provided, the
                dialog loads and edits that profile; otherwise it creates a new
                profile.

        Returns:
            None. Profile data is saved from the nested callback.
        """
        # Treat an empty combobox selection as "create a new profile."
        if existing_profile == "":
            existing_profile = None

        # Create a child dialog for editing profile fields.
        profile_window = tk.Toplevel(self.root)
        profile_window.title("Create/Edit Profile")
        profile_window.geometry("550x450")

        # Local Tkinter variables hold the profile editor's form state.
        profile_name = tk.StringVar(value=existing_profile or "")
        profile_url = tk.StringVar()
        profile_logo_path = tk.StringVar()
        profile_logo_height_percent = tk.StringVar(value="50")
        profile_title_color_hex = tk.StringVar(value="000000")
        profile_heading_color_hex = tk.StringVar(value="000000")
        profile_column_color_hex = tk.StringVar(value="FFFFFF")

        # If an existing profile is being edited, load its saved values into the
        # dialog fields.
        if existing_profile:
            profiles = snapshot_profile_store.load()

            # If the selected profile was deleted or does not exist, close the
            # dialog and tell the user.
            if existing_profile not in profiles:
                messagebox.showerror("Error", "Selected profile does not exist.")
                profile_window.destroy()
                return

            # Copy persisted JSON values into the Tkinter variables.
            profile = profiles[existing_profile]
            profile_url.set(profile["url"])
            profile_logo_path.set(profile["logo_path"])
            profile_logo_height_percent.set(profile["logo_height_percent"])
            profile_title_color_hex.set(profile["title_color_hex"])
            profile_heading_color_hex.set(profile["heading_color_hex"])
            profile_column_color_hex.set(profile["column_color_hex"])

        # Build the profile editor form.
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
            """Persist the profile editor fields to the snapshot profile store.

            Args:
                None. The function reads local Tkinter variables from the
                enclosing profile editor window.

            Returns:
                None. The profile JSON file is updated, or an error dialog is
                shown.
            """
            # Collect the dialog fields into the JSON structure expected by the
            # profile store.
            profile_data = {
                "url": profile_url.get(),
                "logo_path": profile_logo_path.get(),
                "logo_height_percent": profile_logo_height_percent.get(),
                "title_color_hex": profile_title_color_hex.get(),
                "heading_color_hex": profile_heading_color_hex.get(),
                "column_color_hex": profile_column_color_hex.get(),
            }
            try:
                # Save the profile by name, creating or replacing the JSON entry.
                snapshot_profile_store.save(profile_name.get(), profile_data)
            except Exception as exc:
                # If the profile name is missing or saving fails, keep the dialog
                # open and show the error.
                messagebox.showerror("Error", str(exc))
                return

            # Close the editor and refresh the combobox values in the main window.
            profile_window.destroy()
            self.load_profile_names()

        # Save button runs the nested callback above.
        tk.Button(profile_window, text="Save Profile", command=save_new_profile).grid(
            row=7,
            columnspan=3,
            pady=20,
        )

    def load_profile_names(self) -> None:
        """Refresh the report-profile combobox with saved profile names.

        Args:
            None.

        Returns:
            None. The combobox values are replaced in place.
        """
        # Load profile names from JSON and pass them directly to the combobox.
        self.profile_name_combobox["values"] = snapshot_profile_store.names()

    def select_profile_callback(self, _event) -> None:
        """Load the selected profile into the report-generator fields.

        Args:
            _event: Tkinter combobox event object. It is unused because the
                selected profile name is read from ``self.profile_name_var``.

        Returns:
            None. Tkinter variables are updated in place.
        """
        # Load the current profile file so selection uses the latest saved data.
        profiles = snapshot_profile_store.load()
        profile_name = self.profile_name_var.get()

        # If the selected name is no longer present, ignore the event.
        if profile_name not in profiles:
            return

        # Copy stored values into the main form fields.
        profile = profiles[profile_name]
        self.url_var.set(profile["url"])
        self.logo_path_var.set(profile["logo_path"])
        self.logo_height_percent_var.set(profile["logo_height_percent"])
        self.title_color_hex_var.set(profile["title_color_hex"])
        self.heading_color_hex_var.set(profile["heading_color_hex"])
        self.column_color_hex_var.set(profile["column_color_hex"])

    def delete_profile_callback(self) -> None:
        """Delete the selected report-style profile.

        Args:
            None. The selected profile name is read from ``self.profile_name_var``.

        Returns:
            None. A success or error message is shown to the user.
        """
        # If deletion succeeds, refresh the combobox and show confirmation.
        if snapshot_profile_store.delete(self.profile_name_var.get()):
            self.load_profile_names()
            messagebox.showinfo("Success", "Profile deleted successfully.")
        else:
            # If no matching profile existed, tell the user nothing was removed.
            messagebox.showerror("Error", "Profile not found.")


def launch(parent: tk.Tk | tk.Toplevel | None = None):
    """Launch the snapshot report generator workflow window.

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
    SnapshotReportApp(window)

    # Only start a mainloop when this workflow owns the root window.
    if parent is None:
        window.mainloop()
    return window
