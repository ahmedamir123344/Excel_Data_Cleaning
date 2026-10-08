# Excel Data Cleaning — Python (Pandas)

A script that cleans a messy, stacked financial statement
(profit & loss, balance sheet, cash flow) into analysis-ready tables.

## Problem
The source Excel file has three separate financial statements
stacked vertically with inconsistent headers, blank rows, and
duplicate line items. It cannot be analyzed as-is.

## What the script does
- Detects each financial statement section automatically
- Parses year columns from the header row
- Cleans years (FY '09 → 2009) and numeric values
- Handles duplicate line items
- Outputs three wide tables (one per statement)
- Outputs one tidy long table for analysis

## Output files
- `Profit_and_Loss_statement_clean.xlsx`
- `Balance_Sheet_clean.xlsx`
- `Cash_Flow_statement_clean.xlsx`
- `Cola_tidy.xlsx`

## Tools
Python 3.11, Pandas, openpyxl

## Example
Input: 1 messy sheet with 3 stacked statements  
Output: 4 clean, analysis-ready files