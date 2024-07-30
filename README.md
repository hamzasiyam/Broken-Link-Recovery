# Snapshot Analysis Report Generator

This program generates summary and detailed analysis reports for website snapshots. The reports provide insights into the historical captures of a website, helping identify and fix broken links. 

## Features
- Fetches website snapshots from the Internet Archive.
- Extracts date and time components from snapshot links.
- Creates summary and detailed analysis reports in Microsoft Word format.
- Includes the website's logo in the reports.

## Requirements
- Python 3.x
- Required Python packages: `os`, `subprocess`, `pandas`, `datetime`, `python-docx`, `tkinter`
- Waybackpack tool: Install using `pip install waybackpack`
- PIL (Pillow) for image processing: Install using `pip install Pillow`

## Installation
1. Clone the repository:
   ```sh
   git clone https://github.com/hamzasiyam/neverbroken.git
   cd neverbroken
   ```
2. Install required Python packages:
   ```sh
   pip install pandas python-docx pillow waybackpack
   ```

## Usage
### Running the Program
Run the script to open a GUI for generating reports:
```sh
python script_name.py
```

### GUI Components
1. **URL Entry:** Enter the URL of the website to analyze.
2. **Logo File:** Browse and select the logo image file to include in the reports.
3. **Logo Height (%):** Enter the height of the logo as a percentage of its original size.
4. **Generate Reports:** Click to generate the summary and detailed analysis reports.

### Program Functions
#### `get_snapshots(url)`
Fetches the list of snapshots for the given URL using `waybackpack`.

#### `extract_date_from_link(link)`
Extracts the date and time from the snapshot link.

#### `split_date_components(date)`
Splits the date into individual components (month, day, year, hour, minute, second, AM/PM).

#### `create_summary_report(snapshots, domain, logo_path, logo_height_percent, filename='summary_report.docx')`
Creates a summary report of the snapshots with key details.

#### `create_detailed_analysis_report(snapshots, domain, logo_path, logo_height_percent, filename='detailed_analysis_report.docx')`
Creates a detailed analysis report with a table of all snapshots.

#### `browse_file(variable)`
Opens a file dialog to select the logo image file.

#### `generate_reports()`
Generates the summary and detailed analysis reports based on the input URL and logo file.

### Example Output
- **Summary Report:** Provides an overview of the total captures, first capture, and last capture.
- **Detailed Analysis Report:** Includes a table listing all snapshots with date and time details.

## License
This project is licensed under the MIT License.

## Acknowledgements
- Uses the `waybackpack` tool for fetching snapshots.
- Relies on `python-docx` for generating Word documents.
- Utilizes `Pillow` for image processing.
