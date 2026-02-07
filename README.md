# Groupify

Groupify is a small, client-side single-page app that groups student names from an uploaded spreadsheet (Excel or CSV) and lets you export the results. Everything runs in the browser. Perfect for short classroom workflows and quick group assignments.

## Features
- Upload .xlsx/.xls/.csv files entirely client-side.
- Choose which column contains student names.
- Group by number of groups or number of students per group.
- Deterministic (round-robin) or Random (Fisher-Yates) algorithms.
- Optional seed for reproducible random shuffles.
- Export results to XLSX or CSV.

## Quick start
1. Upload `demo/sample_students.xlsx` or your own file.
2. Select the name column, choose grouping mode and algorithm, click "Split groups".
3. Export using the buttons.