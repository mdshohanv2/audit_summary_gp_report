# App Brief: Streamlit Data Analysis App

## Project Overview
This Streamlit application is designed to analyze data from two specific Excel files.

## Features
- **File Upload Section 1:** Named "[Audit report from Shinsa]".
  - Supports Excel (.xlsx, .xls) and CSV (.csv).
  - Target Column: `visit_pos_status`.
  - Action: Count occurrences of each status (case-insensitive, normalized).
  - Action: Handle empty/null values as "unknown/empty".
  - Output: Display summary table with counts and total processed records.
- **File Upload Section 2:** Named "[MTD Successful visit Report]".
  - Supports Excel (.xlsx, .xls) and CSV (.csv).
  - Target Column: `MTD Successful Visits`.
  - Action: Sum all numeric values in the column.
  - Action: Handle non-numeric data gracefully.
  - Output: Display total summation clearly.
