import subprocess
import openpyxl
from datetime import datetime

def run_wget_and_log_to_excel(wget_command, output_excel):
    # Run the wget command and capture the output
    process = subprocess.Popen(wget_command, shell=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE, universal_newlines=True)
    stdout, stderr = process.communicate()
    
    # Create or open the Excel workbook
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.title = "Wget Log"
    
    # Set headers
    sheet['A1'] = "Timestamp"
    sheet['B1'] = "Event"
    sheet['C1'] = "Details"
    
    # Write the wget output to the Excel sheet
    row = 2
    timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    
    if stdout:
        for line in stdout.splitlines():
            sheet[f'A{row}'] = timestamp
            sheet[f'B{row}'] = "Output"
            sheet[f'C{row}'] = line
            row += 1
            
    if stderr:
        for line in stderr.splitlines():
            sheet[f'A{row}'] = timestamp
            sheet[f'B{row}'] = "Error"
            sheet[f'C{row}'] = line
            row += 1
    
    # Save the workbook
    workbook.save(output_excel)
    print(f"Log saved to {output_excel}")

# Define the wget command
wget_command = r'C:\Windows\System32\wget.exe -r -l inf -np -nH --cut-dirs=1 -P ./downloaded_files --convert-links -e robots=off -U Mozilla https://web.archive.org/web/20240814014546/https://hamzasiyam.com/'

# Define the Excel file to save the log
output_excel = "wget_log.xlsx"

# Run the wget command and log to Excel
run_wget_and_log_to_excel(wget_command, output_excel)
