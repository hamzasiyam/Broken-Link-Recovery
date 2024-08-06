import subprocess
import pandas as pd
from datetime import datetime
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import Alignment

def get_snapshots(url):
    # Run the waybackpack command to get the list of snapshots
    result = subprocess.run(['waybackpack', url, '--list'], capture_output=True, text=True)
    if result.returncode != 0:
        print("Error fetching snapshots")
        print(result.stderr)
        return []
    
    # Parse the output to get the links
    snapshots = result.stdout.strip().split('\n')
    return snapshots

def extract_date_from_link(link):
    # Extract date from the Wayback Machine link
    # Example link: http://web.archive.org/web/20230101000000/http://example.com/
    timestamp = link.split('/')[4]
    date = datetime.strptime(timestamp, '%Y%m%d%H%M%S')
    return date

def format_date_columns(date):
    # Format date into separate columns
    return {
        'Month': date.strftime('%B'),
        'Day': date.day,
        'Year': date.year,
        'Hour': date.strftime('%I'),
        'Minute': date.strftime('%M'),
        'Second': date.strftime('%S'),
        'AM/PM': date.strftime('%p')
    }

def save_to_excel(snapshots, filename='snapshots.xlsx'):
    # Prepare data for the first sheet
    data = [(snap, extract_date_from_link(snap)) for snap in snapshots]
    df = pd.DataFrame(data, columns=['Link', 'Date'])
    
    # Create columns for formatted dates
    formatted_dates = df['Date'].apply(format_date_columns).apply(pd.Series)
    
    # Combine the Link, Date and formatted date columns
    df = pd.concat([df[['Link', 'Date']], formatted_dates], axis=1)
    
    # Prepare summary data for the second sheet
    first_capture_date = df['Date'].min()
    last_capture_date = df['Date'].max()
    summary = {
        'Total Captures': [len(snapshots)],
        'First Capture Date': [first_capture_date.strftime('%B %d, %Y %I:%M:%S %p')],
        'Last Capture Date': [last_capture_date.strftime('%B %d, %Y %I:%M:%S %p')]
    }
    summary_df = pd.DataFrame(summary)
    
    # Create an Excel writer
    with pd.ExcelWriter(filename, engine='openpyxl') as writer:
        # Write the data to the first sheet
        df.drop(columns=['Date']).to_excel(writer, sheet_name='Snapshots', index=False)
        
        # Write the summary to the second sheet
        summary_df.to_excel(writer, sheet_name='Summary', index=False)

    # Open the workbook to adjust cell sizes and alignments
    wb = load_workbook(filename)
    ws_snapshots = wb['Snapshots']
    
    # Adjust column width and alignment
    for col in ws_snapshots.columns:
        max_length = 0
        col_letter = get_column_letter(col[0].column)  # Get the column name
        for cell in col:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(cell.value)
            except:
                pass
        adjusted_width = (max_length + 2)
        ws_snapshots.column_dimensions[col_letter].width = adjusted_width
        for cell in col:
            cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)

    ws_summary = wb['Summary']
    for col in ws_summary.columns:
        max_length = 0
        col_letter = get_column_letter(col[0].column)  # Get the column name
        for cell in col:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(cell.value)
            except:
                pass
        adjusted_width = (max_length + 2)
        ws_summary.column_dimensions[col_letter].width = adjusted_width

    # Adjust row heights to default for all rows
    for row in ws_snapshots.iter_rows():
        for cell in row:
            cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        ws_snapshots.row_dimensions[row[0].row].height = None  # Reset to default height

    for row in ws_summary.iter_rows():
        for cell in row:
            cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        ws_summary.row_dimensions[row[0].row].height = None  # Reset to default height

    wb.save(filename)

if __name__ == '__main__':
    url = input("Enter the URL: ")
    snapshots = get_snapshots(url)
    if snapshots:
        save_to_excel(snapshots)
        print(f'Snapshots and summary saved to snapshots.xlsx')
    else:
        print("No snapshots found or there was an error.")
