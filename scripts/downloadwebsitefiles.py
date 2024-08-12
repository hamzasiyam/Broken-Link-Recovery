import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
import subprocess
import os
import time
from datetime import datetime
import urllib.parse
import json

PROFILE_DIR = "./profiles"
PROFILE_FILE = os.path.join(PROFILE_DIR, "proxy_profiles.json")
OUTPUT_DIR = "./downloaded_contents"
wget_process = None  # Global variable to keep track of the wget subprocess

def connect_proxy(proxy_type, domain_name, proxy_port, proxy_username, proxy_password):
    proxy_url = f"{proxy_type}://{proxy_username}:{proxy_password}@{domain_name}:{proxy_port}"
    os.environ['http_proxy'] = proxy_url
    os.environ['https_proxy'] = proxy_url
    print(f"Proxy connected: {proxy_url}")

def disable_proxy():
    os.environ.pop('http_proxy', None)
    os.environ.pop('https_proxy', None)
    print("Proxy disabled")

def print_proxy_info():
    http_proxy = os.environ.get('http_proxy')
    https_proxy = os.environ.get('https_proxy')
    if http_proxy or https_proxy:
        print("Current proxy settings:")
        print(f"HTTP Proxy: {http_proxy}")
        print(f"HTTPS Proxy: {https_proxy}")
    else:
        print("No proxy is currently set.")

def sanitize_filename(filename):
    return "".join(c if c.isalnum() or c in (' ', '.', '_') else '_' for c in filename)

def check_wget():
    try:
        result = subprocess.run(['wget', '--version'], capture_output=True, text=True)
        if result.returncode != 0:
            print("wget is not installed or not found in PATH.")
            return False
        print(f"wget version: {result.stdout.strip()}")
        return True
    except Exception as e:
        print(f"Error checking wget installation: {e}")
        return False

def download_with_wget(url, output_dir):
    global wget_process
    timestamp = datetime.now().strftime("%Y%m%d%H%M%S")
    parsed_url = urllib.parse.urlparse(url)
    domain = parsed_url.netloc
    folder_name = f"{sanitize_filename(domain)}_{timestamp}"
    target_dir = os.path.join(output_dir, folder_name)
    os.makedirs(target_dir, exist_ok=True)
    print(f"Created directory: {target_dir}")
    
    try:
        command = [
            'wget', '--mirror', '--convert-links', '--adjust-extension', '--page-requisites', '--no-parent', url, '-P', target_dir
        ]
        print(f"Running command: {' '.join(command)}")
        wget_process = subprocess.Popen(command, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
        for line in wget_process.stdout:
            print(line.strip())
        wget_process.wait()
        if wget_process.returncode != 0:
            print(f"Error downloading {url}: {wget_process.stderr.read()}")
        else:
            print(f"Successfully downloaded: {url} to {target_dir}")
        return wget_process.returncode == 0
    except Exception as e:
        print(f"Error downloading {url}: {e}")
        return False

def process_spreadsheet(file_path, sheet_name, output_dir, delay):
    try:
        df = pd.read_excel(file_path, sheet_name=sheet_name)
        if 'Capture Link' not in df.columns:
            messagebox.showerror("Error", "'Capture Link' column not found in the spreadsheet.")
            return
        os.makedirs(output_dir, exist_ok=True)
        for i, row in df.iterrows():
            url = row['Capture Link']
            print(f"Downloading URL: {url}")
            success = download_with_wget(url, output_dir)
            if not success:
                print(f"Failed to download: {url}")
            time.sleep(delay)  # Wait for the specified delay before downloading the next archive
        messagebox.showinfo("Success", "Downloaded contents successfully.")
    except Exception as e:
        print(f"Failed to process the spreadsheet: {e}")
        messagebox.showerror("Error", f"Failed to process the spreadsheet: {e}")

def select_file():
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
    if file_path:
        sheet_names = pd.ExcelFile(file_path).sheet_names
        sheet_selector_window(file_path, sheet_names)

def sheet_selector_window(file_path, sheet_names):
    sheet_window = tk.Toplevel(app)
    sheet_window.title("Select Sheet")

    sheet_name_var = tk.StringVar(value=sheet_names[0])
    sheet_label = tk.Label(sheet_window, text="Select Sheet:")
    sheet_label.pack(pady=10)

    sheet_combobox = ttk.Combobox(sheet_window, textvariable=sheet_name_var, values=sheet_names, state="readonly")
    sheet_combobox.pack(pady=10)

    def on_sheet_select():
        selected_sheet = sheet_name_var.get()
        output_dir = os.path.join(os.path.dirname(file_path), "downloaded_contents")
        process_spreadsheet(file_path, selected_sheet, output_dir, int(download_delay_var.get()))
        sheet_window.destroy()

    select_button = tk.Button(sheet_window, text="Select", command=on_sheet_select)
    select_button.pack(pady=20)

def start_program():
    if not check_wget():
        messagebox.showerror("Error", "wget is not installed or not found in PATH.")
        return

    if disable_proxy_var.get() == 1:
        disable_proxy()
    else:
        proxy_type = proxy_type_var.get().lower()
        domain_name = entry_domain_name.get()
        proxy_port = entry_proxy_port.get()
        proxy_username = entry_proxy_username.get()
        proxy_password = entry_proxy_password.get()
        connect_proxy(proxy_type, domain_name, proxy_port, proxy_username, proxy_password)

    print_proxy_info()
    select_file()

def save_profile(name, profile_data):
    if not os.path.exists(PROFILE_DIR):
        os.makedirs(PROFILE_DIR)

    profiles = load_profiles()
    profiles[name] = profile_data

    with open(PROFILE_FILE, "w") as f:
        json.dump(profiles, f, indent=4)

def load_profiles():
    if os.path.exists(PROFILE_FILE):
        with open(PROFILE_FILE, "r") as f:
            return json.load(f)
    return {}

def delete_profile(profile_name):
    profiles = load_profiles()
    if profile_name in profiles:
        del profiles[profile_name]
        with open(PROFILE_FILE, "w") as f:
            json.dump(profiles, f, indent=4)

def select_profile(profile_name_var, vars):
    profiles = load_profiles()
    profile_name = profile_name_var.get()
    if profile_name in profiles:
        profile = profiles[profile_name]
        vars['proxy_type'].set(profile["proxy_type"])
        vars['domain_name'].delete(0, tk.END)
        vars['domain_name'].insert(0, profile["domain_name"])
        vars['proxy_port'].delete(0, tk.END)
        vars['proxy_port'].insert(0, profile["proxy_port"])
        vars['proxy_username'].delete(0, tk.END)
        vars['proxy_username'].insert(0, profile["proxy_username"])
        vars['proxy_password'].delete(0, tk.END)
        vars['proxy_password'].insert(0, profile["proxy_password"])

def create_profile_window():
    profile_window = tk.Toplevel(app)
    profile_window.title("Create Profile")

    profile_name = tk.StringVar()
    profile_proxy_type = tk.StringVar(value="HTTP")
    profile_domain_name = tk.StringVar()
    profile_proxy_port = tk.StringVar()
    profile_proxy_username = tk.StringVar()
    profile_proxy_password = tk.StringVar()

    tk.Label(profile_window, text="Profile Name:").grid(row=0, column=0, pady=10, sticky=tk.W)
    tk.Entry(profile_window, textvariable=profile_name, width=50).grid(row=0, column=1)

    tk.Label(profile_window, text="Proxy Type:").grid(row=1, column=0, pady=10, sticky=tk.W)
    ttk.Combobox(profile_window, textvariable=profile_proxy_type, values=["HTTP", "SOCKS5"], state="readonly").grid(row=1, column=1)

    tk.Label(profile_window, text="Domain Name:").grid(row=2, column=0, pady=10, sticky=tk.W)
    tk.Entry(profile_window, textvariable=profile_domain_name, width=50).grid(row=2, column=1)

    tk.Label(profile_window, text="Proxy Port:").grid(row=3, column=0, pady=10, sticky=tk.W)
    tk.Entry(profile_window, textvariable=profile_proxy_port, width=50).grid(row=3, column=1)

    tk.Label(profile_window, text="Proxy Username:").grid(row=4, column=0, pady=10, sticky=tk.W)
    tk.Entry(profile_window, textvariable=profile_proxy_username, width=50).grid(row=4, column=1)

    tk.Label(profile_window, text="Proxy Password:").grid(row=5, column=0, pady=10, sticky=tk.W)
    tk.Entry(profile_window, textvariable=profile_proxy_password, width=50, show="*").grid(row=5, column=1)

    def save_new_profile():
        profile_data = {
            "proxy_type": profile_proxy_type.get(),
            "domain_name": profile_domain_name.get(),
            "proxy_port": profile_proxy_port.get(),
            "proxy_username": profile_proxy_username.get(),
            "proxy_password": profile_proxy_password.get()
        }
        save_profile(profile_name.get(), profile_data)
        profile_window.destroy()
        load_profile_names()

    tk.Button(profile_window, text="Save Profile", command=save_new_profile).grid(row=6, columnspan=2, pady=20)

def edit_profile_window():
    selected_profile = profile_name_var.get()
    profiles = load_profiles()
    if selected_profile not in profiles:
        messagebox.showerror("Error", "Selected profile does not exist.")
        return

    profile_window = tk.Toplevel(app)
    profile_window.title("Edit Profile")

    profile_name = tk.StringVar(value=selected_profile)
    profile_proxy_type = tk.StringVar(value=profiles[selected_profile]["proxy_type"])
    profile_domain_name = tk.StringVar(value=profiles[selected_profile]["domain_name"])
    profile_proxy_port = tk.StringVar(value=profiles[selected_profile]["proxy_port"])
    profile_proxy_username = tk.StringVar(value=profiles[selected_profile]["proxy_username"])
    profile_proxy_password = tk.StringVar(value=profiles[selected_profile]["proxy_password"])

    tk.Label(profile_window, text="Profile Name:").grid(row=0, column=0, pady=10, sticky=tk.W)
    tk.Entry(profile_window, textvariable=profile_name, width=50, state="readonly").grid(row=0, column=1)

    tk.Label(profile_window, text="Proxy Type:").grid(row=1, column=0, pady=10, sticky=tk.W)
    ttk.Combobox(profile_window, textvariable=profile_proxy_type, values=["HTTP", "SOCKS5"], state="readonly").grid(row=1, column=1)

    tk.Label(profile_window, text="Domain Name:").grid(row=2, column=0, pady=10, sticky=tk.W)
    tk.Entry(profile_window, textvariable=profile_domain_name, width=50).grid(row=2, column=1)

    tk.Label(profile_window, text="Proxy Port:").grid(row=3, column=0, pady=10, sticky=tk.W)
    tk.Entry(profile_window, textvariable=profile_proxy_port, width=50).grid(row=3, column=1)

    tk.Label(profile_window, text="Proxy Username:").grid(row=4, column=0, pady=10, sticky=tk.W)
    tk.Entry(profile_window, textvariable=profile_proxy_username, width=50).grid(row=4, column=1)

    tk.Label(profile_window, text="Proxy Password:").grid(row=5, column=0, pady=10, sticky=tk.W)
    tk.Entry(profile_window, textvariable=profile_proxy_password, width=50, show="*").grid(row=5, column=1)

    def save_edited_profile():
        profile_data = {
            "proxy_type": profile_proxy_type.get(),
            "domain_name": profile_domain_name.get(),
            "proxy_port": profile_proxy_port.get(),
            "proxy_username": profile_proxy_username.get(),
            "proxy_password": profile_proxy_password.get()
        }
        save_profile(profile_name.get(), profile_data)
        profile_window.destroy()
        load_profile_names()

    tk.Button(profile_window, text="Save Profile", command=save_edited_profile).grid(row=6, columnspan=2, pady=20)

def delete_selected_profile():
    selected_profile = profile_name_var.get()
    if messagebox.askyesno("Delete Profile", f"Are you sure you want to delete the profile '{selected_profile}'?"):
        delete_profile(selected_profile)
        load_profile_names()
        clear_profile_fields()

def clear_profile_fields():
    proxy_type_var.set("HTTP")
    entry_domain_name.delete(0, tk.END)
    entry_proxy_port.delete(0, tk.END)
    entry_proxy_username.delete(0, tk.END)
    entry_proxy_password.delete(0, tk.END)

def load_profile_names():
    profiles = load_profiles()
    profile_names = list(profiles.keys())
    profile_name_combobox['values'] = profile_names

def select_profile_callback(event):
    select_profile(profile_name_var, {
        'proxy_type': proxy_type_var,
        'domain_name': entry_domain_name,
        'proxy_port': entry_proxy_port,
        'proxy_username': entry_proxy_username,
        'proxy_password': entry_proxy_password
    })

def on_closing():
    global wget_process
    if wget_process is not None:
        print("Terminating wget process...")
        wget_process.terminate()
        wget_process = None
    app.destroy()

app = tk.Tk()
app.title("Web Archive Downloader with Proxy")

frame = tk.Frame(app, padx=10, pady=10)
frame.pack(padx=10, pady=10)

# Summary of what the program does
summary_label = tk.Label(frame, text="This program downloads web content based on URLs from an Excel spreadsheet, "
                                      "supports proxy settings, and saves the downloaded content into organized directories.",
                         wraplength=400, justify="left", font=("Arial", 10, "italic"))
summary_label.grid(row=0, columnspan=3, pady=10)

tk.Label(frame, text="Proxy Type:").grid(row=1, column=0, pady=5, sticky=tk.W)
proxy_type_var = tk.StringVar(value="HTTP")
proxy_type_menu = ttk.Combobox(frame, textvariable=proxy_type_var, values=["HTTP", "SOCKS5"], state="readonly")
proxy_type_menu.grid(row=1, column=1, pady=5)

tk.Label(frame, text="Domain Name:").grid(row=2, column=0, pady=5, sticky=tk.W)
entry_domain_name = tk.Entry(frame, width=50)
entry_domain_name.grid(row=2, column=1, pady=5)

tk.Label(frame, text="Proxy Port:").grid(row=3, column=0, pady=5, sticky=tk.W)
entry_proxy_port = tk.Entry(frame, width=50)
entry_proxy_port.grid(row=3, column=1, pady=5)

tk.Label(frame, text="Proxy Username:").grid(row=4, column=0, pady=5, sticky=tk.W)
entry_proxy_username = tk.Entry(frame, width=50)
entry_proxy_username.grid(row=4, column=1, pady=5)

tk.Label(frame, text="Proxy Password:").grid(row=5, column=0, pady=5, sticky=tk.W)
entry_proxy_password = tk.Entry(frame, width=50, show="*")
entry_proxy_password.grid(row=5, column=1, pady=5)

disable_proxy_var = tk.IntVar()
disable_proxy_checkbutton = tk.Checkbutton(frame, text="Disable Proxy", variable=disable_proxy_var)
disable_proxy_checkbutton.grid(row=6, columnspan=2, pady=5)

tk.Label(frame, text="Time between downloads (seconds):").grid(row=7, column=0, pady=5, sticky=tk.W)
download_delay_var = tk.StringVar(value="5")
tk.Entry(frame, textvariable=download_delay_var, width=10).grid(row=7, column=1, pady=5, sticky=tk.W)

button_start = tk.Button(frame, text="Start", command=start_program)
button_start.grid(row=8, columnspan=2, pady=10)

tk.Label(frame, text="Profile Name").grid(row=9, column=0, pady=10, sticky=tk.W)
profile_name_var = tk.StringVar()
profile_name_combobox = ttk.Combobox(frame, textvariable=profile_name_var, state='readonly')
profile_name_combobox.grid(row=9, column=1, pady=10, padx=10)
profile_name_combobox.bind("<<ComboboxSelected>>", select_profile_callback)

button_create_profile = tk.Button(frame, text="Create Profile", command=create_profile_window)
button_create_profile.grid(row=9, column=2, pady=10)

button_edit_profile = tk.Button(frame, text="Edit Profile", command=edit_profile_window)
button_edit_profile.grid(row=10, column=1, pady=10)

button_delete_profile = tk.Button(frame, text="Delete Profile", command=delete_selected_profile)
button_delete_profile.grid(row=10, column=2, pady=10)

load_profile_names()

app.protocol("WM_DELETE_WINDOW", on_closing)
app.mainloop()
