# Meat Company Sales Report System

Windows desktop application for reading daily customer-order Excel files and generating Traditional Chinese daily and monthly sales reports in Excel and PDF.

## Features

- Reads separate Excel files for each customer order.
- Calculates revenue from sold quantity and the selected F, G, or H selling price.
- Generates daily Excel and one-page PDF reports.
- Generates monthly Excel and two-page PDF reports with comparisons and trends.
- Detects duplicate orders, invalid filenames, missing fields, and date mismatches.
- Automatically selects and creates dated order folders.
- Provides a Traditional Chinese desktop interface for nontechnical staff.

## Requirements

- Windows 11
- Python 3.12
- Dependencies from `requirements.txt`

## Development Setup

```powershell
python -m pip install -r requirements.txt
python scripts/preflight_check.py
python scripts/daily_report_app.py
```

## Main Folders

- `Orders`: daily order files. Customer data is ignored by Git.
- `Daily Reports`: generated daily reports and tracked order template.
- `Monthly Reports`: generated monthly reports.
- `scripts`: application, report reader, PDF generator, and preflight checks.
- `config`: portable relative-path settings.
- `assets`: application icon.

Order filenames use `YYYYMMDD_客戶名_序號.xlsx`.

## Data Privacy

Real order files and generated reports must not be committed. `.gitignore` excludes these folders while retaining the blank folder structure and order template.
