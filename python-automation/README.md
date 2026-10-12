# Sales Report Automation

One command turns a raw sales CSV into a finished, formatted Excel report. It replaces the manual routine of cleaning a file, building summary tables and drawing a chart.

```bash
pip install pandas openpyxl
python auto_report.py sample_input.csv Sales_Report.xlsx
```

## What it does
1. Reads the CSV and stops with a clear message if a required column is missing.
2. Cleans the data and records every fix: duplicate rows, inconsistent names, missing unit counts, negative revenue, unreadable dates.
3. Summarises revenue by month, branch and category.
4. Writes one workbook with six sheets: **Summary**, **By month** (with a native Excel chart), **By branch**, **By category**, **Cleaning log** and **Clean data** (filterable).

## Sample run
`sample_input.csv` is the simulated 2025 supermarket sales file from the `lagos-retail-analysis` project (13,165 rows). Real output:

```
Read       13,165 rows from sample_input.csv
Fixed          25  exact duplicate rows: removed
Fixed          10  inconsistent branch names (spacing or capitals): standardised
Fixed          20  missing unit counts: rebuilt from revenue and price (20), removed (0)
Fixed          10  negative revenue values: sign corrected
Clean      13,140 rows
Wrote     Sales_Report.xlsx  (6 sheets, 1 chart) in 0.8s
```

## Checked
The report's total revenue (₦2,362,697,570.99) and units sold (685,763) match the separate analysis in `lagos-retail-analysis` to the kobo, and the monthly figures match.

## Limits
- It expects the column names `date, branch, category, units_sold, unit_price_ngn, discount, revenue_ngn`. A different file layout needs the column list at the top of the script changed.
- The report shows values, not live Excel formulas.
- The sample file is simulated, not real company data.
