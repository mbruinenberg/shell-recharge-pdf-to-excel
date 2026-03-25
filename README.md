# ShellRechargePy

A Python utility that extracts charging session data from Shell Recharge PDF invoices and generates a consolidated Excel summary for bookkeeping and administration.

## What it does

Shell Recharge sends individual PDF receipts for each EV charging session. This tool processes a folder of those PDFs and produces a single Excel spreadsheet containing all session details, ready for financial administration.

**Extracted data per session:**
- Receipt number and issue date
- Charging start/end date, time, and duration
- Station address and city
- Energy consumed (kWh) and price per kWh
- Transaction fee
- Amount excl. VAT, VAT %, VAT amount, amount incl. VAT
- Payment method

**Excel output features:**
- Formatted headers with auto-filter and frozen header row
- EUR currency formatting
- Totals row with sum formulas for energy, costs, and VAT
- Records sorted chronologically by session start date

## Requirements

- Python 3.13+
- [pdfplumber](https://github.com/jsvine/pdfplumber) - PDF text extraction
- [openpyxl](https://openpyxl.readthedocs.io/) - Excel file generation

## Setup

```bash
# Create and activate virtual environment
python -m venv .venv
. .venv/Scripts/activate      # Windows (Git Bash)
# or: source .venv/bin/activate  # Linux/macOS

# Install dependencies
pip install pdfplumber openpyxl
```

## Usage

```bash
python shell_recharge_extractor.py <pdf_folder> [output_file.xlsx]
```

**Arguments:**
| Argument | Required | Description |
|---|---|---|
| `pdf_folder` | Yes | Path to folder containing Shell Recharge PDF invoices |
| `output_file.xlsx` | No | Output filename. Defaults to `Shell_Recharge_Summary_YYYY-MM.xlsx` in the PDF folder |

**Examples:**

```bash
# Process all PDFs in the folder, auto-generate output filename
python shell_recharge_extractor.py "./invoices"

# Specify a custom output file
python shell_recharge_extractor.py "./invoices" "March_2026.xlsx"
```

**Output:**
```
Processing: receipt_001.pdf
Processing: receipt_002.pdf
Processing: receipt_003.pdf
Done! 3 receipt(s) -> Shell_Recharge_Summary_2026-03.xlsx
```

## Automation

A Windows batch script (`ShellRechargeMonthlyrun.bat`) is included for monthly automation. It processes multiple charging folders and archives the PDFs after extraction.

## License

This project is for personal/internal use.
