# CSV vs XLSX Hostname Comparison Tool

A user-friendly web application built with Flask to compare hostnames between CSV and XLSX files, finding unique hostnames from the CSV file that don't exist in the XLSX file.

## Features

- **Web-Based Interface**: Access through any web browser
- **Drag and Drop Interface**: Easily drag and drop CSV and XLSX files
- **Browse Option**: Click to browse and select files
- **Column Selection**: Specify which columns to compare (defaults to Column A/Index 0)
- **Real-time Results**: See comparison results immediately
- **Export Functionality**: Download unique hostnames as a CSV file
- **Modern UI**: Beautiful, responsive design

## Installation

1. Install Python 3.7 or higher
2. Install required packages:
```bash
pip install -r requirements.txt
```

## Usage

1. Start the Flask server:
```bash
python app.py
```

2. Open your web browser and navigate to:
```
http://localhost:5000
```

3. **Select Files**:
   - Drag and drop your CSV file (e.g., `host_status_13.csv`) into the CSV drop zone, OR
   - Click the CSV drop zone to browse and select your CSV file
   - Drag and drop your XLSX file (e.g., `all.xlsx`) into the XLSX drop zone, OR
   - Click the XLSX drop zone to browse and select your XLSX file

4. **Configure Columns** (Optional):
   - By default, the tool compares Column A (Index 0) from both files
   - You can specify different columns using:
     - Numeric index (0, 1, 2, etc.)
     - Excel column letters (A, B, C, etc.)

5. **Compare**: Click the "Compare Files" button

6. **View Results**: The unique hostnames will be displayed with statistics

7. **Export** (Optional): Click "Export Results to CSV" to download the unique hostnames

## How It Works

This tool replicates the Excel VLOOKUP functionality:
- It reads hostnames from the CSV file (source)
- It reads hostnames from the XLSX file (reference)
- It finds hostnames in CSV that don't exist in XLSX (equivalent to filtering #N/A results from VLOOKUP)
- It displays and allows export of these unique hostnames

## Example

If you have:
- **CSV file** (`host_status_13.csv`) with hostnames: server1, server2, server3, server4
- **XLSX file** (`all.xlsx`) with hostnames: server1, server3, server5

The tool will find:
- **Unique hostnames**: server2, server4 (present in CSV but not in XLSX)

## Requirements

- Python 3.7+
- pandas
- openpyxl
- flask
- werkzeug

## Notes

- The comparison is case-sensitive
- Empty cells are automatically excluded
- Hostnames are trimmed of whitespace before comparison
- Duplicate hostnames within each file are handled automatically
- Maximum file size: 50MB per file
- The server runs on `http://localhost:5000` by default

