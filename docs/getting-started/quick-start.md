# Quick Start

Get started with XLEX in 5 minutes.

## Your First Workbook

Create a new Excel file:

```bash
xlex create sales.xlsx
```

Check the file info:

```bash
xlex info sales.xlsx
# Sheets: Sheet1
# Total Cells: 0
```

## Working with Cells

Set individual cell values:

```bash
# Set text
xlex cell set sales.xlsx Sheet1 A1 "Product"
xlex cell set sales.xlsx Sheet1 B1 "Revenue"

# Set numbers
xlex cell set sales.xlsx Sheet1 A2 "Widget"
xlex cell set sales.xlsx Sheet1 B2 1500

# Set formula
xlex cell set sales.xlsx Sheet1 B5 "=SUM(B2:B4)"
```

Read cell values:

```bash
xlex cell get sales.xlsx Sheet1 A1
# Product

xlex cell get sales.xlsx Sheet1 B5
# =SUM(B2:B4)
```

## Adding Rows

Append rows quickly:

```bash
# Single row (comma-separated)
xlex row append sales.xlsx Sheet1 "Gadget,2500"
xlex row append sales.xlsx Sheet1 "Gizmo,1800"

# Multiple values in row
xlex row append sales.xlsx Sheet1 "Total,=SUM(B2:B4)"
```

## Export Data

Export to different formats:

```bash
# Export to CSV
xlex export csv sales.xlsx sales.csv

# Export to JSON
xlex export json sales.xlsx sales.json --header

# Export to stdout
xlex export csv sales.xlsx -
```

## Import Data

Import from external files:

```bash
# Import CSV
xlex import csv data.csv spreadsheet.xlsx

# Import JSON
xlex import json records.json spreadsheet.xlsx
```

## Working with Sheets

```bash
# List sheets
xlex sheet list workbook.xlsx

# Add a new sheet
xlex sheet add workbook.xlsx "Analysis"

# Rename a sheet
xlex sheet rename workbook.xlsx Sheet1 Data

# Copy a sheet
xlex sheet copy workbook.xlsx Data "Data Backup"
```

## Pipeline Examples

XLEX works great with Unix pipelines:

```bash
# Create from CSV data
echo "Name,Age
Alice,30
Bob,25" | xlex import csv - output.xlsx

# Extract and process
xlex export csv input.xlsx - | grep "2024" | wc -l

# Chain commands
xlex export json data.xlsx - | jq '.[] | select(.status=="active")'
```

## Template Processing

Use templates for report generation:

```bash
# Create a template with placeholders
# In template.xlsx: {{name}}, {{date}}, {{#each items}}...{{/each}}

# Apply template with JSON data
xlex template apply template.xlsx data.json report.xlsx
```

## Configuration

Create a project config file:

```bash
xlex config init
# Creates .xlex.yml
```

Example `.xlex.yml`:

```yaml
defaults:
  format: json
  sheet: Data
  
export:
  csv:
    delimiter: ";"
    
style:
  header:
    bold: true
    background: "#4472C4"
```

## Next Steps

- Learn about [Sheet Operations](../commands/sheet.md)
- Explore [Cell Commands](../commands/cell.md)
- Read about [Template Syntax](../reference/template-syntax.md)
- Check [Pipeline Integration](../guides/pipelines.md)
