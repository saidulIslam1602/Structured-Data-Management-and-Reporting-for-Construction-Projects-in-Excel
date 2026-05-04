# Structured Data Management and Reporting for Construction Projects in Excel

> A professional, formula-driven Excel workbook for construction project coordination — built with Python (`openpyxl`) and covering every aspect of site data management, progress tracking, risk control, and monthly reporting.

---

## Overview

This project delivers a fully automated Excel reporting system designed for a **Project Coordinator / Data Analyst** role in the construction sector. The workbook is generated programmatically via a Python script, ensuring reproducibility and easy customisation.

It models real-world construction site workflows:
- Weekly discipline progress monitoring
- Risk identification and scoring
- Action item tracking and escalation
- NCR (Non-Conformance Report) quality control
- Monthly Earned Value and schedule performance reporting
- Centralised data validation and audit logging

---

## Workbook Structure — 12 Sheets

| # | Sheet | Purpose |
|---|-------|---------|
| 1 | **DASHBOARD** | Executive KPI overview — live COUNTIF/SUMIFS formulas, 2 embedded charts, hyperlinks |
| 2 | **Weekly Progress** | 18 activities across 8 disciplines — XLOOKUP, Variance%, Status badges |
| 3 | **Risk Register** | 15 risks — P×I score formula, IFS-driven Level, heatmap conditional formatting |
| 4 | **Action Item Log** | 17 actions — VLOOKUP for owner, Days Overdue formula (updates on open) |
| 5 | **NCR Quality Tracker** | 12 NCRs — INDEX/MATCH for discipline lead, Days Open counter |
| 6 | **Monthly Report** | 16 entries across 3 months — SUMIFS aggregation, SV%, CV%, CPI chart |
| 7 | **Data Validation Log** | Audit trail of data quality issues found and resolved |
| 8 | **Submission Tracker** | Per-discipline weekly data submission status (Submitted / Pending) |
| 9 | **Meeting Log** | Coordination meeting minutes, decisions, and action item links |
| 10 | **Lookup Tables** | Master reference: disciplines, leads, risk matrix, status definitions |
| 11 | **Power BI Export** | Flat denormalised table — import directly into Power BI via Get Data → Excel |
| 12 | **Instructions & Guide** | Step-by-step usage guide for each sheet |

---

## Advanced Excel Capabilities

### Formulas Used
| Formula | Where Used | Purpose |
|---------|-----------|---------|
| `COUNTIF` / `COUNTIFS` | Dashboard | Count activities/risks by status |
| `SUMIFS` | Dashboard, Monthly Report | Aggregate hours and EV/PV by filter |
| `AVERAGEIF` / `AVERAGEIFS` | Dashboard | Average actual % by discipline |
| `XLOOKUP` | Weekly Progress | Auto-fill responsible person from Lookup Tables |
| `VLOOKUP` | Action Item Log | Look up discipline lead by discipline name |
| `INDEX / MATCH` | NCR Tracker, Monthly Report | Flexible cross-table lookup |
| `IFS` | Risk Register | Multi-tier risk level classification |
| `IF / IFERROR` | All sheets | Conditional logic and safe error handling |
| `TODAY()` | Action Log, NCR Tracker | Live overdue and days-open counters |

### Power BI Integration
A dedicated **Power BI Export** sheet (`tblPowerBI_Export`) contains a single flat, denormalised table combining data from all three main trackers (Weekly Progress, Risk Register, Action Item Log) with consistent column names typed for direct import:

```
Power BI Desktop → Home → Get Data → Excel Workbook → select file → check tblPowerBI_Export → Load
```

Suggested Power BI visuals: use `Source_Sheet` as a slicer, `Discipline` as bar-chart axis, `Status` and `Risk_Level` as legends, `Planned_Pct` / `Actual_Pct` as measures.

### Excel Tables (Native Filter & Sort)
All data sheets are wrapped in proper **Excel Table** objects:

```
tblWeeklyProgress   tblRiskRegister     tblActionLog
tblNCRTracker       tblMonthlyReport    tblValidationLog
tblSubmissionTracker  tblMeetingLog     tblPowerBI_Export
```

Each table provides:
- **Filter dropdown** on every column header — click `▼` to filter by value, colour, or text
- **Sort A→Z / Z→A** on every column — ascending, descending, or custom sort
- **Auto-expansion** — add a row at the bottom and the table grows automatically
- **Structured references** for clean formula syntax

### Data Validation
Every input column carries typed validation with a **hover tooltip**:

| Column Type | Validation Rule | Example Error Message |
|-------------|----------------|----------------------|
| Discipline | Dropdown from Lookup Tables | "Choose from the list" |
| Progress % | Decimal 0.0 – 1.0 | "Enter 0.75 for 75%, not 75" |
| Probability / Impact | Integer 1 – 5 | "Enter a whole number 1–5" |
| Week Number | Integer 1 – 52 | "Enter ISO week 1–52" |
| All date columns | Date > 01-Jan-2020 | "Enter DD-MMM-YY format" |
| Status / Priority / Type | Dropdown list | Custom lists per sheet |

### Conditional Formatting
| Sheet | Rule | Visual Effect |
|-------|------|---------------|
| Weekly Progress | Variance % | Red → White → Green colour scale |
| Weekly Progress | Status badges | Muted pastel: Completed=green, Behind=red, Ahead=blue |
| Risk Register | Risk Score | Red → Amber → Green heatmap (P×I matrix) |
| Risk Register | Risk Level badges | High=red, Medium=amber, Low=green |
| Action Item Log | Days Overdue | Red fill when > 0; amber data bar |
| NCR Tracker | Days Open | Amber > 14 days; red > 28 days |
| Monthly Report | SV% / CV% | Green positive, red negative |
| Dashboard | KPI cells | Accent-left border, live formula values |

### Row Grouping — Monthly Report
The Monthly Report uses Excel's **Outline / Group** feature. Three month blocks (Feb, Mar, Apr) each have a `–` collapse button on the left row margin — click to hide detail rows and see only the month summary, `+` to expand.

### Charts (5 Embedded)
| Sheet | Chart | Type |
|-------|-------|------|
| DASHBOARD | Discipline Progress: Planned vs Actual % | Clustered column |
| DASHBOARD | Schedule Performance Index by Discipline | Horizontal bar |
| Weekly Progress | Activity Progress: Planned vs Actual | Horizontal bar |
| Monthly Report | Earned Value vs Planned Value by Month | Clustered column |
| Monthly Report | CPI Trend Over Time | Line chart |

### Column Header Comments
Every formula column carries a **hover comment** (small red triangle in corner) explaining:
- The exact formula logic
- Why the cell must not be manually edited
- How it connects to other sheets

### Hyperlinks
Discipline names on the **DASHBOARD** are clickable — they jump directly to `Weekly Progress!A3`.

### Print Setup
- **Print Titles**: Header row repeats on every printed page (all sheets)
- **Freeze Panes**: Row 1–2 frozen (title + header visible while scrolling)
- **Landscape orientation** on wide sheets
- **Fit-to-page** printing

### Named Ranges
```
DisciplineLeadTable  →  'Lookup Tables'!$A$6:$B$14
DisciplineList       →  'Lookup Tables'!$A$6:$A$14
```

---

## File Structure

```
.
├── build_excel.py                           # Python generator script (openpyxl)
├── Construction_Project_Data_Management.xlsx  # Generated Excel workbook
└── README.md                               # This file
```

---

## Requirements

```
Python 3.8+
openpyxl >= 3.1.0
```

Install dependency:
```bash
pip install openpyxl
```

---

## Usage

### Generate the workbook
```bash
python build_excel.py
```

The script builds all 9 sheets, adds charts and advanced features, and saves the `.xlsx` file to the same directory.

### Open in Excel or LibreOffice
Open `Construction_Project_Data_Management.xlsx` in:
- **Microsoft Excel 2016+** — full feature support (recommended)
- **LibreOffice Calc 7+** — fully compatible; press **F9** to recalculate formulas on first open

### Data Entry Workflow
1. **Lookup Tables** — verify or update discipline leads and reference lists first
2. **Weekly Progress** — enter `Planned %` and `Actual %` (cols G & H) each week; all other columns auto-calculate
3. **Risk Register** — enter `Probability` (1–5) and `Impact` (1–5); Score and Level auto-calculate
4. **Action Item Log** — enter action details; `Days Overdue` updates automatically on every open
5. **NCR Quality Tracker** — log NCRs; `Days Open` updates daily
6. **Monthly Report** — enter budgeted/actual hours and EV/PV values; SV% and CV% auto-calculate
7. **DASHBOARD** — all KPIs and charts refresh automatically from the above data

### How to Filter Data
1. Click the **▼ dropdown arrow** in any column header on a data sheet
2. Choose **Filter by Value**, **Filter by Colour**, or type in the search box
3. Click **OK** — only matching rows are shown; all formulas still calculate correctly
4. To clear: click the arrow again → **Clear Filter**

### How to Sort Data
1. Click the **▼ dropdown arrow** on the column you want to sort
2. Choose **Sort A → Z**, **Sort Z → A**, or **Sort by Cell Colour**
3. For multi-level sort: **Data → Sort** (Excel) or right-click → Sort

---

## Dashboard KPIs

| KPI | Formula Logic |
|-----|---------------|
| Total Activities | `COUNTA(Weekly Progress!A4:A21)` |
| Completed | `COUNTIF(Weekly Progress!K4:K21, "Completed")` |
| Behind Schedule | `COUNTIF(Weekly Progress!K4:K21, "Behind")` |
| Avg Actual Progress | `AVERAGEIF(Weekly Progress!H4:H21, "<>", ...)` |
| Open Risks | `COUNTIF(Risk Register!M4:M18, "Open")` |
| High Risks | `COUNTIF(Risk Register!H4:H18, "High")` |
| Open Actions | `COUNTIFS(Action Log!I4:I20, "<>Closed")` |
| Overdue Actions | `COUNTIF(Action Log!N4:N20, ">0")` |
| Open NCRs | `COUNTIF(NCR Tracker!L4:L15, "Open")` |

---

## Formulas Reference Card

### Risk Score & Level (Risk Register)
```excel
Score  = E4 * F4                              (Probability × Impact)
Level  = IFS(G4>=15, "High", G4>=8, "Medium", G4<8, "Low")
```

### Activity Status (Weekly Progress)
```excel
Status = IF(H4>=1, "Completed",
           IF(H4-G4>0.03, "Ahead",
             IF(G4-H4>0.03, "Behind", "On Track")))
```

### Discipline Lead Lookup (XLOOKUP)
```excel
= IFERROR(XLOOKUP(B4, LookupTables!A:A, LookupTables!B:B), "—")
```

### Days Overdue (Action Item Log)
```excel
= IF(AND(I4<>"Closed", ISNUMBER(G4), TODAY()>G4), TODAY()-G4, "")
```

### Schedule Variance % (Monthly Report)
```excel
= IFERROR((F4-E4)/E4, 0)      (Actual% - Planned%) / Planned%
```

### Cost Variance % (Monthly Report)
```excel
= IFERROR((G4-H4)/H4, 0)      (EV - PV) / PV
```

---

## Customisation

| What to change | Where |
|----------------|-------|
| Add disciplines | `Lookup Tables` sheet → Column A rows 6–14 |
| Change discipline leads | `Lookup Tables` sheet → Column B rows 6–14 |
| Add more activities | `Weekly Progress` — add rows inside the Table (auto-expands) |
| Change date range | `Monthly Report` — edit Column A (month labels) |
| Re-generate from scratch | Edit `build_excel.py` → run `python build_excel.py` |

---

## Skills Demonstrated

This project showcases the core Excel and data analysis skills required for a Project Coordinator / Data Analyst role in the AEC (Architecture, Engineering & Construction) sector:

- **Data Structuring** — multi-sheet relational data model with lookup tables as master reference
- **Advanced Formulas** — XLOOKUP, INDEX/MATCH, COUNTIFS, SUMIFS, IFS, IFERROR, TODAY()
- **Dynamic Reporting** — dashboard KPIs that update automatically as data changes
- **Conditional Formatting** — heatmaps, colour scales, data bars, formula-based badge rules
- **Data Validation** — dropdown lists, range constraints, date validation, input messages
- **Excel Tables** — structured tables enabling native sort/filter/auto-expand
- **Data Visualisation** — 5 embedded charts (column, bar, line) across 3 sheets
- **Earned Value Management** — SV%, CV%, CPI trend tracking
- **Risk Management** — probability × impact matrix, automated risk scoring
- **Quality Control** — NCR tracking with automated days-open escalation
- **Python Automation** — entire workbook generated programmatically with openpyxl

---

## License

MIT — free to use, adapt, and share with attribution.
