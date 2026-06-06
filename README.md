# Tableau to Power BI Migration Tool

A lightweight Tableau workbook analysis tool designed to accelerate Tableau-to-Power BI migrations.

The tool parses Tableau workbooks (.twb and .twbx) and automatically generates migration documentation that identifies:

- Dashboards
- Worksheets
- Fields
- Calculations
- Dependencies
- Dashboard-to-worksheet relationships
- Calculation references
- Migration complexity metrics

The output is an Excel workbook that can be used as a migration roadmap and technical reference during Power BI development.

---

## Features

### Migration Workbook Generator

Creates:

- Dashboard Map
- Migration Roadmap
- Calculation Dictionary
- Migration Summary
- Calculation Dependencies

### Workbook Enhancement

Automatically formats generated workbooks by:

- Creating Excel tables
- Auto-sizing columns
- Freezing headers
- Applying filters
- Improving readability

### Tableau Support

Supported formats:

- .twb
- .twbx

---

## Requirements

Python 3.9+

Required package:

```bash
pip install openpyxl
