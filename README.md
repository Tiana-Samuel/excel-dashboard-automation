# Excel Dashboard Automation

**One-click margin analysis dashboard** – automating 34 manual Excel steps with a Script Lab snippet.

## Project Objective

Building a product-level margin dashboard manually in Excel requires:

- Writing lots of complex formulas (`SUMIFS`, `SUMPRODUCT`, `INDEX/MATCH`, nested `IF`)
- Applying number formatting (currency, percentages)
- Setting up conditional formatting (green/yellow/red fills)
- Creating a combo chart (clustered columns + line on secondary axis)
I automated the entire process so that a dashboard can be generated from raw transaction data in under a second eliminating manual work

## What I Built

A Script Lab snippet (TypeScript + Office.js) that:

1. Reads raw transaction data (year, quarter, product, revenue, cost, profit)
2. Writes dynamic Excel formulas into a `Dashboard` sheet
3. Formats cells and applies conditional formatting
4. Creates a summary table for charting
5. Inserts and configures a professional combo chart

The script performs all 34 steps listed in the changelog below – automatically.

## The Changelog 

| Step | Manual Step (now automated) |
|------|-------------------------------|
| 1 | `SUMIFS` formula for Total Revenue |
| 2 | Format as currency, 0 decimals |
| 3 | Copy formula down 31 rows |
| 4 | `SUMPRODUCT` formula for Weighted Average Margin |
| 5 | Format as percentage, 1 decimal |
| 6 | Copy formula down |
| 7 | `IF` formula for Rolling 3-Month Trend |
| 8 | Format as percentage, 1 decimal |
| 9 | Copy formula down |
| 10 | `IF/INDEX/MATCH` formula for YoY Margin Delta |
| 11 | Format as percentage, 1 decimal |
| 12 | Copy formula down |
| 13 | Nested `IF` for Margin Health (Strong/Moderate/At Risk) |
| 14 | Copy formula down |
| 15 | Conditional formatting: green / yellow / red fills |
| 16 | Add section label "Chart Data" |
| 17 | Add column headers (Quarter, Widget Pro, etc.) |
| 18 | Add quarter labels (2023 Q1 to 2024 Q4) |
| 19-22 | `SUMPRODUCT` formulas for each product’s margin data |
| 23 | `SUMIF` formula for Total Revenue by quarter |
| 24 | Copy formulas down |
| 25-26 | Format percentages and currency |
| 27-28 | Insert clustered column chart |
| 29-30 | Move Total Revenue to secondary axis, change to line chart |
| 31-33 | Add chart title and axis titles |
| 34 | Position chart below data tables |

> **Instead of doing these 34 steps manually for each new dataset, the script executes them all in one click.**

## Formula Logic 

### 1. Total Revenue per Product/Quarter
=SUMIFS('Raw Data'!E:E, 'Raw Data'!D:D, A8, 'Raw Data'!B:B, LEFT(B8,4), 'Raw Data'!C:C, RIGHT(B8,2))

Matches product (column A), extracts year and quarter from column B, sums revenue.

### 2. Weighted Average Margin
=SUMPRODUCT(('Raw Data'!D:D=A8) * ('Raw Data'!B:B&" "&'Raw Data'!C:C = B8) * ('Raw Data'!G:G)) / C8

Multiplies condition arrays with margin values, divides by total revenue.

### 3. Rolling Quarter-over-Quarter Trend
=IF(B8="2023 Q1", "N/A", D8 - D7)

Shows margin change from previous quarter; first quarter returns “N/A”.

### 4. Year-over-Year Margin Delta
=IF(LEFT(B8,4)="2023", "N/A", D8 - INDEX(D$8:D$39, MATCH(1, (A$8:A$39=A8) * (B$8:B$39=(LEFT(B8,4)-1)&RIGHT(B8,3)), 0)))

Compares current quarter margin to same quarter previous year using `INDEX/MATCH` with multiple criteria.

### 5. Margin Health Classification
=IF(D8>0.35, "Strong", IF(D8>=0.2, "Moderate", "At Risk"))

Thresholds: >35% Strong, 20–35% Moderate, <20% At Risk.

### 6. Chart Summary Formulas
- **Product margins:** `SUMPRODUCT` matching product name and quarter
- **Total revenue:** `SUMIF` summing revenue for each quarter across all products

## Automation Benefits

| Manual approach | Automated script |
|----------------|------------------|
| 10–15 minutes per dashboard | < 1 second |
| High risk of cell reference errors | Zero errors |
| Must re-do for each new dataset | Works instantly on any properly structured data |
| Chart needs manual reconfiguration | Chart updates dynamically |

## Raw Data Structure Expected

The script works with an Excel sheet named `Raw Data` containing:

- **Column B** – Year (e.g., 2023)
- **Column C** – Quarter (Q1, Q2, Q3, Q4)
- **Column D** – Product name
- **Column E** – Revenue (numeric)
- **Column G** – Margin (decimal: Profit / Revenue)

If only Profit is available, Margin is calculated as `Profit / Revenue` before running the script.

## Output Dashboard Features

After execution, the `Dashboard` sheet contains:

| Range | Content |
|-------|---------|
| `C8:C39` | Total revenue (currency, zero decimals) |
| `D8:D39` | Weighted avg margin (percentage, 1 decimal) |
| `E8:E39` | Rolling trend (or N/A for first quarter) |
| `F8:F39` | YoY delta (or N/A for 2023) |
| `G8:G39` | Health status with color coding |
| `A43:F43` | Chart column headers |
| `A44:A51` | Quarters 2023 Q1 – 2024 Q4 |
| `B44:E51` | Product margins for chart |
| `F44:F51` | Total revenue per quarter |
| `A53:H75` | Combo chart (columns + line on secondary axis) |

## Technology Stack

- **Excel** (Office.js API)
- **Script Lab** (snippet runner)
- **TypeScript** (formula injection, conditional formatting, chart creation)

## Repository Contents

- `margin-dashboard.json` – the complete Script Lab snippet
- `README.md` – this project documentation
