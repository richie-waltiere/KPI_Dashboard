# KPI Dashboard
Excel-based KPI dashboard prototype designed for monitoring manufacturing-style data using Excel formulas, conditional formatting, and VBA automation.

## Features

### Dashboard Metrics
The dashboard automatically summarizes:
- Total planned units
- Total actual units
- Overall production efficiency
- Total defects
- Total downtime
- Critical production alerts

### Operator Summary
Generates operator-level summaries including:
- Assigned production lines
- Planned vs actual production totals
- Overall efficiency
- Defect counts
- Downtime totals

### Conditional Formatting
Conditional formatting is used to highlight abnormal operational conditions.

#### Efficiency Thresholds
- Efficiency >= 97% → Green
- Efficiency >= 94% and < 97% → Yellow
- Efficiency < 94% → Red

#### Operational Alerts
- Total defects > 50 → Red
- Total downtime > 200 minutes → Red

### VBA Automation — "Generate Alert Report"
The "Generate Alert Report" button runs a VBA macro that:
1. Identifies rows marked as "Critical"
2. Creates a new worksheet named "Alert Report"
3. Copies all critical rows into the generated report

### Charts
The dashboard includes visual summaries for:
- Defects by production line
- Downtime by production line
- Defects by product
- Downtime by product

## Technologies Used
- Microsoft Excel
- VBA
- Dynamic Excel formulas
- Conditional formatting
- Structured table references

## Purpose
This project was created as a prototype dashboard focused on KPI reporting for a manufacturing environment, abnormal-case identification, and workflow automation.
