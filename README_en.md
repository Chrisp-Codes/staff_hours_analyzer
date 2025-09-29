# Shift Hour Analysis Tool
![Python](https://img.shields.io/badge/python-3.10%2B-blue)
![Status](https://img.shields.io/badge/status-POC-orange)
![License: MIT](https://img.shields.io/badge/License-MIT-green.svg)
[![English](https://img.shields.io/badge/README-English-informational?style=flat-square)](README_en.md)
[![Deutsch](https://img.shields.io/badge/README-Deutsch-informational?style=flat-square)](README.md)

## Purpose

This tool analyzes employee working hours and calculates staff hours per time slot. It's useful for evaluating workforce productivity — e.g. in the hospitality industry.

## Features

- Reads an Excel export (expects a specific format in this version)
- Splits shifts into hourly segments
- Rounds to two decimal places
- Groups by date and hour
- Auto-formatting (centering, column width)
- Exports a cleaned Excel file with a new name
- GUI for file selection (no code contact required)

## Export Format

- Uses the sheet named “Alle Mitarbeiter”
- Header starts at row 7
- Required columns: Tag, Startzeit, Endzeit, Dauer netto (dezimal)
- Stops processing when column A contains the string “Summe:”

## Usage

1. Install Python
2. Install dependencies:

```bash
pip install pandas openpyxl
```

3. Run the tool:

```bash
python gui.py
```

## Test Data

Test files are located in the `example_data/` folder

## Planned

- Configurable column mappings (for generic exports)
- Productivity KPIs (revenue per hour/employee)
- Monthly/weekly evaluations with median comparisons
- Portable `.exe` for non-technical users
