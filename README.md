# Data Cleaning & Reporting Automation

An automated Python pipeline that ingests raw, messy sales data, performs standardized data preprocessing, and generates formatted Excel reports, visual charts, and an HTML summary dashboard.

## Key Features
- **Automated Preprocessing**:
  - Trims extra whitespace and standardizes text casing (`Customer_Name`, `Region`, `Category`, `Product`, `Payment_Method`).
  - Fixes inconsistent city spellings (`Banglore` -> `Bengaluru`).
  - Parses mixed date formats (`YYYY-MM-DD`, `DD/MM/YYYY`, `Jan 08 2026`) into a unified `YYYY-MM-DD` format.
  - Removes exact and duplicate `Order_ID` records.
  - Imputes missing numeric values (`Quantity`, `Unit_Price`) using category-level medians and fills missing categorical labels cleanly.
- **Automated Reporting & Visual Summaries**:
  - `output/cleaned_sales_data.csv`: Cleaned dataset ready for analysis.
  - `output/automated_sales_report.xlsx`: Multi-sheet Excel workbook (`Cleaned_Data`, `Category_Summary`, `Region_Summary`, `Data_Quality_Audit`).
  - `output/visual_summary.png`: 4-panel dashboard chart (Category Revenue, Region Revenue, Daily Trend, Cleaning Audit).
  - `output/automation_report.html`: Standalone summary report combining audit metrics, tables, and charts.

## Project Structure
```text
data-cleaning-reporting-automation/
├── data/
│   └── raw_sales_data.csv
├── output/                  # Generated after running the script
├── clean_and_report.py
├── requirements.txt
└── README.md
```

## How to Run

### Option 1: Local Python
```bash
pip install -r requirements.txt
python clean_and_report.py
```

### Option 2: Google Colab (No Local Setup Needed)
1. Upload `data/raw_sales_data.csv` and `clean_and_report.py` to Colab (keeping `raw_sales_data.csv` inside a `data/` folder).
2. Run `!python clean_and_report.py`.
3. Download the generated files from the `output/` folder.
