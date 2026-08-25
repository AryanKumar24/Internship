# Data Cleaning & Reporting Automation

A Python pipeline that automates the full workflow: raw messy data → cleaned
data → automated report with charts. Built for the Thiranex internship task.

## What it does

1. **Load** — reads a raw CSV/Excel file.
2. **Clean**
   - Standardizes text (trims whitespace, fixes casing like `"north"` vs `"NORTH"` vs `" East "`)
   - Parses inconsistent date formats (`2025-01-05`, `05/01/2025`, `01-05-2025`) into one consistent format
   - Removes exact duplicate rows
   - Flags and handles invalid values (e.g. negative units sold)
   - Fills missing values — median for numeric columns, mode for categorical columns
   - Logs every cleaning action taken (see `outputs/cleaning_log.txt`)
3. **Analyze** — builds summary statistics: totals, averages, revenue by region/product/rep, monthly trend.
4. **Visualize** — generates bar charts and a trend line chart (`outputs/charts/`).
5. **Report** — assembles everything into a formatted, multi-sheet Excel workbook (`outputs/sales_report.xlsx`):
   - **Summary** sheet: KPIs + embedded charts
   - **Breakdowns** sheet: revenue tables by region/product/rep
   - **Cleaned Data** sheet: the full cleaned dataset

## How to run

```bash
pip install pandas numpy matplotlib openpyxl
python clean_and_report.py
```

Outputs land in `outputs/`:
- `cleaned_sales_data.csv`
- `cleaning_log.txt`
- `charts/*.png`
- `sales_report.xlsx`

## Using your own data

Replace `data/raw_sales_data.csv` with your real dataset (or point
`RAW_DATA_PATH` in `clean_and_report.py` at it). The included
`generate_sample_data.py` only exists to create realistic messy sample data —
delete it once you're using real data. If your column names differ, update
the column references in `clean_data()` and `build_summary()` to match.

## Key concepts demonstrated

- **Data preprocessing**: missing-value imputation strategy, deduplication,
  format standardization, and outlier/invalid-value handling
- **Automation**: a single script takes raw input to finished report with no
  manual steps, and is re-runnable on any new raw file of the same shape
- **Reporting efficiency**: one script produces both machine-readable output
  (cleaned CSV) and a stakeholder-ready deliverable (formatted Excel report
  with visuals), instead of doing cleaning and reporting as separate manual steps
