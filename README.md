<p align="center">
  <img src="banner.png" alt="Sales & Financial Performance Dashboard Banner" width="100%">
</p>

<h1 align="center">📊 Sales & Financial Performance Dashboard</h1>
<p align="center">
  Advanced Excel analytics project simulating end-to-end retail sales analysis — from raw transactions to an interactive KPI dashboard.
</p>

<p align="center">
  <img src="https://img.shields.io/badge/Excel-Advanced-217346?style=flat-square&logo=microsoft-excel&logoColor=white">
  <img src="https://img.shields.io/badge/Python-Data%20Generation-3776AB?style=flat-square&logo=python&logoColor=white">
  <img src="https://img.shields.io/badge/Formulas-10%2C275-1F4E78?style=flat-square">
  <img src="https://img.shields.io/badge/Errors-0-63BE7B?style=flat-square">
  <img src="https://img.shields.io/badge/License-MIT-yellow?style=flat-square">
</p>

---

## 📋 Table of Contents

- [Overview](#overview)
- [Dashboard Preview](#dashboard-preview)
- [Key Features](#key-features)
- [Tech & Skills Used](#tech--skills-used)
- [File Structure](#file-structure)
- [How to Use](#how-to-use)
- [Extending This Project Natively in Excel](#extending-this-project-natively-in-excel)
- [License](#license)

---

## Overview

This project simulates **2 years (FY2024–FY2025)** of retail sales across **4 regions**, **4 product categories**, and **10 salespeople**, then builds a fully formula-driven Excel workbook on top of it — the kind of deliverable a Data Analyst produces when asked to "make sense of the sales data" for stakeholders.

> **Note on tooling:** Python was used only to generate the synthetic raw dataset (2,527 transactions). Every calculation in the workbook — revenue, cost, profit, pivot summaries, scenario outputs, KPIs — is a **live Excel formula**, not a precomputed value. Change any input and the entire workbook recalculates.

## Dashboard Preview

<p align="center">
  <img src="Sales Excel dashboard_screenshot.png" alt="Dashboard Preview" width="90%">
</p>

## Key Features

| Sheet | What it demonstrates |
|---|---|
| **Dashboard** | KPI cards, an interactive Region filter (data validation dropdown), and 4 linked charts (trend, category mix, top performers, region comparison) |
| **Raw_Sales_Data** | 2,527-row transactional dataset with formula-driven Revenue / Cost / Profit / Margin columns |
| **Formulas_Reference** | `SUMIFS`, `XLOOKUP` (with `INDEX/MATCH` fallback for older Excel), `INDEX-MATCH`, nested `IF` commission-tier logic |
| **Pivot_Analysis** | Region × Month revenue matrix, category performance summary, and ranked salesperson leaderboard — all `SUMIFS`-driven so they update automatically |
| **What_If_Analysis** | Scenario Manager-style Best / Likely / Worst case inputs, plus a discount % sensitivity table (Goal Seek-ready) |

## Tech & Skills Used

- **Excel formulas:** SUMIFS, XLOOKUP, INDEX-MATCH, nested IF, RANK, IFERROR
- **Data validation:** dropdown-driven scenario and region filters
- **Conditional formatting:** color scales, icon sets, KPI-style highlighting
- **Charts:** line, bar, pie — all linked to formula-driven tables
- **What-if analysis:** scenario inputs, sensitivity table (Goal Seek-compatible)
- **Python:** synthetic dataset generation (`pandas`/`csv`, `random`)

## File Structure

```
├── README.md                              # This file
├── banner.png                             # Project banner
├── synthetic_sales_data.csv               # Raw synthetic dataset (2,527 transactions)
├── Sales_Performance_Workbook.xlsx        # Full analytical workbook (6 sheets)
├── Sales Performance_Excel Dashboard.xlsx # Clean, dashboard-only version (1 sheet)
├── Sales Excel dashboard_screenshot.png   # Preview of the dashboard
└── Sales & Financial Performance Excel Dashboard.pdf  # Project Report
└── LICENSE                                            # MIT License
```

## How to Use

**Full analytical workbook** —
 `Sales_Performance_Workbook.xlsx`
1. Download and open in Excel (2019+ recommended for XLOOKUP support; falls back gracefully on older versions)
2. Start on the **Dashboard** tab — use the Region dropdown to filter
3. Explore **What_If_Analysis** to test Best/Likely/Worst scenarios
4. Optional: follow the **README** tab inside the workbook to add Power Query, Power Pivot, DAX, and a VBA macro

**Single-sheet dashboard** — 
 `Sales Performance_Excel Dashboard.xlsx`
1. Download and open — everything lives on one clean **Dashboard** sheet
2. Use the Region filter dropdown to update KPIs and charts instantly
3. A quick, presentation-ready view for anyone who just wants the dashboard without the underlying analysis sheets

## Extending This Project Natively in Excel

The README tab inside the workbook includes exact steps to add:
- **Power Query** — import/clean `Raw_Sales_Data` via Get Data
- **Power Pivot / Data Model** — sample DAX measures (Total Revenue, YoY Growth %)
- **Slicers** — for Region, Category, Year
- **VBA macro** — one-click refresh + PDF export of the Dashboard

## License

This project is licensed under the **MIT License** — see the [LICENSE](LICENSE) file for details.
```

