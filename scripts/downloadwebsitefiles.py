import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
import subprocess
import os

def connect_proxy(proxy_type, domain_name, proxy_port, proxy_username, proxy_password):
    proxy_url = f"{proxy_type}://{proxy_username}:{proxy_password}@{domain_name}:{proxy_port}"
    os.environ['http_proxy'] = proxy_url
    os.environ['https_proxy'] = proxy_url

def download_with_waybackpack(url, output_dir):
    try:
        command = ['waybackpack', url, '-d', output_dir, '--ignore-errors', '--quiet']
        result = subprocess.run(command, capture_output=True, text=True)
        if result.returncode != 0:
            print(f"Error downloading {url}: {result.stderr}")
        return result.returncode == 0
    except Exception as e:
        print(f"Error downloading {url}: {e}")
        return False

def process_spreadsheet(file_path, output_dir):
    try:
        df = pd.read_excel(file_path)
        if 'Capture Link' not in df.columns:
            messagebox.showerror("Error", "'Capture Link' column not found in the spreadsheet.")
            return
        os.makedirs(output_dir, exist_ok=True)
        for i, row in df.iterrows():
            url = row['Capture Link']
            success = download_with_waybackpack(url, output_dir)
            if not success:
                print(f"Failed to download: {url}")
        messagebox.showinfo("Success", "Downloaded contents successfully.")
    except Exception as e:
        messagebox.showerror("Error", f"Failed to process the spreadsheet: {e}")

def select_file():
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
    if file_path:
        output_dir = os.path.join(os.path.dirname(file_path), "downloaded_contents")
        process_spreadsheet(file_path, output_dir)

def start_program():
    proxy_type = proxy_type_var.get().lower()
    domain_name = entry_domain_name.get()
    proxy_port = entry_proxy_port.get()
    proxy_username = entry_proxy_username.get()
    proxy_password = entry_proxy_password.get()
    connect_proxy(proxy_type, domain_name, proxy_port, proxy_username, proxy_password)
    select_file()

app = tk.Tk()
app.title("Web Archive Downloader with Proxy")

frame = tk.Frame(app, padx=10, pady=10)
frame.pack(padx=10, pady=10)

tk.Label(frame, text="Proxy Type:").grid(row=0, column=0, pady=5, sticky=tk.W)
proxy_type_var = tk.StringVar(value="HTTP")
proxy_type_menu = ttk.Combobox(frame, textvariable=proxy_type_var, values=["HTTP", "SOCKS5"], state="readonly")
proxy_type_menu.grid(row=0, column=1, pady=5)

tk.Label(frame, text="Domain Name:").grid(row=1, column=0, pady=5, sticky=tk.W)
entry_domain_name = tk.Entry(frame, width=50)
entry_domain_name.grid(row=1, column=1, pady=5)

tk.Label(frame, text="Proxy Port:").grid(row=2, column=0, pady=5, sticky=tk.W)
entry_proxy_port = tk.Entry(frame, width=50)
entry_proxy_port.grid(row=2, column=1, pady=5)

tk.Label(frame, text="Proxy Username:").grid(row=3, column=0, pady=5, sticky=tk.W)
entry_proxy_username = tk.Entry(frame, width=50)
entry_proxy_username.grid(row=3, column=1, pady=5)

tk.Label(frame, text="Proxy Password:").grid(row=4, column=0, pady=5, sticky=tk.W)
entry_proxy_password = tk.Entry(frame, width=50, show="*")
entry_proxy_password.grid(row=4, column=1, pady=5)

button_start = tk.Button(frame, text="Start", command=start_program)
button_start.grid(row=5, columnspan=2, pady=10)

app.mainloop()
