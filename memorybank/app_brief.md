# App Brief: Audit & MTD Report Analyzer

## Project Overview
A professional, minimal Streamlit application designed for the GP Audit Team to analyze and compare Audit Reports (Shinsa) with MTD Successful Visit Reports.

## Features & Logic
- **UI Design:** Minimal "shadcn-like" aesthetic with a wide layout for laptop optimization.
- **Reporting Period:** Automatically extracted from the `To Date` column in the MTD report. Formats as `DD-Month-YYYY` (e.g., 13-February-2026).
- **Theme Support:** Fully compatible with both light and dark modes with theme-aware styling.

### Data Source 1: Audit Report (Shinsa)
- **File Types:** Excel (.xlsx, .xls) and CSV.
- **Target Column:** `visit_pos_status`.
- **Logic:** Normalized status counts (Open, Temporarily Closed, Permanently Closed, Moved, Not Found).
- **Handling:** Null values treated as "unknown/empty".

### Data Source 2: MTD Successful Visit Report
- **File Types:** Excel (.xlsx, .xls) and CSV.
- **Target Columns:** 
    - `MTD Successful Visits` (Total summation).
    - `Visit Status Open`, `Visit Status Temporary Closed`, etc. (Individual sums).
    - `To Date` (For report period).
- **Handling:** Robust matching for multi-line headers (new-line characters) and varying cases.

### Final Summary Report Table
Combines data from both sources with the following calculations:
1. **Successful Visits:** Summed from MTD Report per category.
2. **Visit Ach%:** `(Successful Visits of Row) / (Total Successful Visits)`.
3. **Audited Visits:** Counted from Shinsa Report per category.
4. **Audit Ach%:** `(Audited Visits of Row) / (Total Audited Visits)`.
5. **Coverage %:** `(Audited Visits of Row) / (Successful Visits of Row)`.
6. **Grand Totals:** Aggregated sums and 100.00% benchmark for achievement columns.

### Export Features
- **Excel Export:** A stylised Excel file with:
    - Custom dark header with the report title.
    - Auto-adjusted column widths.
    - Centered cell alignments.
    - Bolded Grand Total row.
    - Download button positioned cleanly at the bottom-left under the final coverage summary.
