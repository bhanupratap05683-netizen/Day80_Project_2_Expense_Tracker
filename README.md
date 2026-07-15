# Project 2: Expense Tracker — Day 80 Build

## Overview
A production-oriented expense tracking system that ingests raw transaction data, categorizes expenses via configurable regex rules, compares actual spending against predefined budgets, and exports a multi-sheet Excel report with embedded visualizations.

## Day 80 Scope
- Regex-based expense categorization engine
- Budget variance analysis (Actual vs. Budget vs. Variance %)
- Multi-sheet Excel report architecture (Raw Data | Processed Data | Summary Dashboard)
- Automated matplotlib chart generation and embedding via openpyxl
- Execution logging and input validation

## Input
- `day80_input.xlsx`: Auto-generated practice file containing raw transactions and budget limits.

## Output
- `day80_expense_report.xlsx`: Formatted, self-contained workbook with embedded charts
- `day80_charts/`: Exported PNG charts for external use
- `expense_tracker_day80.log`: Execution and insight log

## Quick Start
```bash
python day80_Project 2 - Expense Tracker.py
