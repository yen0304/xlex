# xlex Examples

Real-world patterns for common Excel automation tasks. Each example shows a complete workflow, not just a single command.

## Table of Contents

- [Build a styled financial report](#build-a-styled-financial-report)
- [Process CSV data into a formatted workbook](#process-csv-data-into-a-formatted-workbook)
- [Batch generate invoices from template](#batch-generate-invoices-from-template)
- [Data pipeline with unix tools](#data-pipeline-with-unix-tools)
- [Audit and fix formulas](#audit-and-fix-formulas)
- [Multi-sheet dashboard](#multi-sheet-dashboard)
- [Safe editing workflow](#safe-editing-workflow)
- [Bulk data entry](#bulk-data-entry)

---

## Build a styled financial report

Create a professional-looking report from scratch with headers, data, formulas, and styling.

```bash
# 1. Create workbook with two sheets
xlex create report.xlsx --sheets Summary,Details

# 2. Set up headers
xlex cell set report.xlsx Summary A1 "Metric"
xlex cell set report.xlsx Summary B1 "Q1"
xlex cell set report.xlsx Summary C1 "Q2"
xlex cell set report.xlsx Summary D1 "Q3"
xlex cell set report.xlsx Summary E1 "Q4"
xlex cell set report.xlsx Summary F1 "Total"

# 3. Style the header row
xlex range style report.xlsx Summary A1:F1 --bold --bg-color 4472C4 --text-color FFFFFF --font-size 12
xlex range border report.xlsx Summary A1:F1 --style medium --border-color 2F5496

# 4. Add data
xlex row append report.xlsx Summary "Revenue,150000,165000,180000,195000"
xlex row append report.xlsx Summary "COGS,90000,95000,100000,105000"
xlex row append report.xlsx Summary "Gross Profit,,,,"
xlex row append report.xlsx Summary "Operating Expenses,35000,38000,40000,42000"
xlex row append report.xlsx Summary "Net Income,,,,"

# 5. Add formulas
# Gross Profit = Revenue - COGS
xlex formula set report.xlsx Summary B4 "B2-B3"
xlex formula set report.xlsx Summary C4 "C2-C3"
xlex formula set report.xlsx Summary D4 "D2-D3"
xlex formula set report.xlsx Summary E4 "E2-E3"

# Net Income = Gross Profit - OpEx
xlex formula set report.xlsx Summary B6 "B4-B5"
xlex formula set report.xlsx Summary C6 "C4-C5"
xlex formula set report.xlsx Summary D6 "D4-D5"
xlex formula set report.xlsx Summary E6 "E4-E5"

# Total column = sum of quarters
xlex formula set report.xlsx Summary F2 "SUM(B2:E2)"
xlex formula set report.xlsx Summary F3 "SUM(B3:E3)"
xlex formula set report.xlsx Summary F4 "SUM(B4:E4)"
xlex formula set report.xlsx Summary F5 "SUM(B5:E5)"
xlex formula set report.xlsx Summary F6 "SUM(B6:E6)"

# 6. Column widths and freeze
xlex column width report.xlsx Summary A 20.0
xlex column width report.xlsx Summary B 12.0
xlex column width report.xlsx Summary C 12.0
xlex column width report.xlsx Summary D 12.0
xlex column width report.xlsx Summary E 12.0
xlex column width report.xlsx Summary F 14.0
xlex style freeze report.xlsx Summary --rows 1 --cols 1

# 7. Conditional formatting on Net Income
xlex style condition report.xlsx Summary B6:E6 --highlight-cells --gt 0 --bg-color C6EFCE
xlex style condition report.xlsx Summary B6:E6 --highlight-cells --lt 0 --bg-color FFC7CE

# 8. Verify
xlex validate report.xlsx
```

## Process CSV data into a formatted workbook

Take raw CSV data, import it, and add structure.

```bash
# 1. Import
xlex import csv sales_data.csv sales.xlsx --header -s RawData

# 2. Check what we have
xlex range get sales.xlsx RawData A1:Z1 -f json    # read headers
xlex stats sales.xlsx                               # row count

# 3. Style the header row
xlex range style sales.xlsx RawData A1:F1 --bold --bg-color 305496 --text-color FFFFFF

# 4. Add auto-filter feel with freeze
xlex style freeze sales.xlsx RawData --rows 1

# 5. Add a summary sheet
xlex sheet add sales.xlsx Summary -p 0              # insert at position 0 (first)
xlex cell set sales.xlsx Summary A1 "Sales Summary"
xlex cell set sales.xlsx Summary A3 "Total Records"
xlex formula calc count sales.xlsx RawData A2:A1000  # get count, then set it
xlex cell set sales.xlsx Summary A4 "Total Revenue"
xlex formula calc sum sales.xlsx RawData D2:D1000    # get sum

# 6. Export a clean version
xlex export markdown sales.xlsx - -s RawData | head -20   # preview in chat
```

## Batch generate invoices from template

Generate one invoice per customer from a template file and a JSON data source.

```bash
# 1. Inspect the template
xlex template list invoice_template.xlsx

# 2. Validate against data
xlex template validate invoice_template.xlsx --vars customers.json

# 3. Preview one
xlex template preview invoice_template.xlsx --vars customers.json

# 4. Generate all invoices
xlex template apply invoice_template.xlsx invoices/ \
  --vars customers.json \
  --per-record \
  --output-pattern "invoice_{index}.xlsx"
```

Where `customers.json` might look like:
```json
[
  {"name": "Acme Corp", "amount": 15000, "date": "2026-03-01", "invoice_id": "INV-001"},
  {"name": "Globex Inc", "amount": 8500, "date": "2026-03-01", "invoice_id": "INV-002"}
]
```

## Data pipeline with unix tools

Combine xlex with standard unix tools for data transformation.

```bash
# Filter rows where column 3 (C) > 1000, keep as xlsx
xlex export csv data.xlsx - -s Sales | \
  awk -F, 'NR==1 || $3 > 1000' | \
  xlex import csv /dev/stdin filtered.xlsx --header

# Count rows per unique value in column A
xlex export csv data.xlsx - -s Data | \
  tail -n +2 | \
  cut -d, -f1 | \
  sort | uniq -c | sort -rn

# Merge two sheets into one CSV for external processing
xlex export csv report.xlsx - -s Q1 > /tmp/combined.csv
xlex export csv report.xlsx - -s Q2 | tail -n +2 >> /tmp/combined.csv
xlex import csv /tmp/combined.csv combined.xlsx --header

# JSON processing with jq
xlex export json data.xlsx - -s Sheet1 --header | \
  jq '[.[] | select(.status == "active")]' > active_records.json
```

## Audit and fix formulas

Check a workbook for formula issues and fix them.

```bash
# 1. Find all formulas
xlex formula list data.xlsx Sheet1

# 2. Check for errors
xlex formula check data.xlsx

# 3. Detect circular references
xlex formula circular data.xlsx

# 4. Find what depends on a cell
xlex formula refs data.xlsx Sheet1 B1 --dependents

# 5. Find what a cell depends on
xlex formula refs data.xlsx Sheet1 D10 --precedents

# 6. Bulk replace sheet references (e.g., after rename)
xlex formula replace data.xlsx Sheet1 "OldSheet!" "NewSheet!"

# 7. Get formula statistics
xlex formula stats data.xlsx
```

## Multi-sheet dashboard

Create a workbook with multiple sheets that reference each other.

```bash
# Create structure
xlex create dashboard.xlsx --sheets Overview,Sales,Costs,Inventory

# Populate Sales sheet
xlex cell set dashboard.xlsx Sales A1 "Month"
xlex cell set dashboard.xlsx Sales B1 "Revenue"
xlex row append dashboard.xlsx Sales "Jan,50000"
xlex row append dashboard.xlsx Sales "Feb,55000"
xlex row append dashboard.xlsx Sales "Mar,62000"

# Populate Costs sheet
xlex cell set dashboard.xlsx Costs A1 "Month"
xlex cell set dashboard.xlsx Costs B1 "Expenses"
xlex row append dashboard.xlsx Costs "Jan,30000"
xlex row append dashboard.xlsx Costs "Feb,32000"
xlex row append dashboard.xlsx Costs "Mar,35000"

# Overview sheet references other sheets
xlex cell set dashboard.xlsx Overview A1 "Dashboard"
xlex range style dashboard.xlsx Overview A1:A1 --bold --font-size 16

xlex cell set dashboard.xlsx Overview A3 "Total Revenue"
xlex formula set dashboard.xlsx Overview B3 "SUM(Sales!B2:B100)"

xlex cell set dashboard.xlsx Overview A4 "Total Costs"
xlex formula set dashboard.xlsx Overview B4 "SUM(Costs!B2:B100)"

xlex cell set dashboard.xlsx Overview A5 "Net Profit"
xlex formula set dashboard.xlsx Overview B5 "B3-B4"

# Style the overview
xlex range style dashboard.xlsx Overview A3:A5 --bold
xlex column width dashboard.xlsx Overview A 20.0
xlex column width dashboard.xlsx Overview B 15.0

# Define named ranges for clarity
xlex range name dashboard.xlsx TotalRevenue "Overview!B3"
xlex range name dashboard.xlsx TotalCosts "Overview!B4"
```

## Safe editing workflow

When modifying important files, use a non-destructive approach.

```bash
# 1. Validate the original
xlex validate important.xlsx

# 2. Clone it first
xlex clone important.xlsx backup_$(date +%Y%m%d).xlsx

# 3. Make changes with --dry-run to preview
xlex cell set important.xlsx Sheet1 A1 "New Value" --dry-run

# 4. Apply changes
xlex cell set important.xlsx Sheet1 A1 "New Value"

# 5. Validate after changes
xlex validate important.xlsx

# 6. Compare with original
xlex export json important.xlsx - -s Sheet1 --header > new.json
xlex export json backup_*.xlsx - -s Sheet1 --header > old.json
diff old.json new.json
```

## Bulk data entry

Efficiently load many values into a workbook.

```bash
# Option 1: Row append (comma-separated)
for i in $(seq 1 100); do
  xlex row append data.xlsx Sheet1 "Item $i,$((RANDOM % 1000)),$((RANDOM % 50))"
done

# Option 2: Batch mode from JSON (more efficient for large updates)
cat <<'EOF' | xlex cell batch data.xlsx
{"sheet": "Sheet1", "cell": "A1", "value": "Name", "type": "string"}
{"sheet": "Sheet1", "cell": "B1", "value": "Score", "type": "string"}
{"sheet": "Sheet1", "cell": "A2", "value": "Alice", "type": "string"}
{"sheet": "Sheet1", "cell": "B2", "value": "95", "type": "number"}
{"sheet": "Sheet1", "cell": "A3", "value": "Bob", "type": "string"}
{"sheet": "Sheet1", "cell": "B3", "value": "87", "type": "number"}
EOF

# Option 3: Import from existing data file
xlex import json records.json data.xlsx -s Imported
```
