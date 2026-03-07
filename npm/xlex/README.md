# xlex

<p align="center">
  <img src="https://raw.githubusercontent.com/yen0304/xlex/main/logo.png" alt="xlex logo" width="640">
</p>

<p align="center">
  <strong>A streaming CLI Excel manipulation tool for developers and automation pipelines.</strong>
</p>

## Installation

```bash
npm install -g xlex
```

Or use directly with npx:

```bash
npx xlex info report.xlsx
```

## Usage

```bash
# Display workbook information
xlex info report.xlsx

# Get a cell value
xlex cell get report.xlsx Sheet1 A1

# Set a cell value
xlex cell set report.xlsx Sheet1 A1 "Hello, World!"

# Export to CSV
xlex export csv report.xlsx -s Sheet1 > data.csv

# Import from JSON
xlex import json data.json output.xlsx

# Search across all sheets (like Ctrl+F in Excel)
xlex search report.xlsx "revenue"
xlex search report.xlsx "error" -s Sheet1          # restrict to one sheet
xlex search report.xlsx "^2026-" -r                # regex search
xlex search report.xlsx "total" -c B -f json       # column filter + JSON output
```

### More Commands

```bash
# Sheets
xlex sheet list report.xlsx
xlex sheet add report.xlsx NewSheet
xlex sheet rename report.xlsx OldName NewName

# Rows & Columns
xlex row append data.xlsx Sheet1 "a,b,c"
xlex row find data.xlsx Sheet1 "pattern"
xlex column width data.xlsx Sheet1 A 20.0

# Ranges
xlex range get data.xlsx Sheet1 A1:D10 -f json
xlex range fill data.xlsx Sheet1 A1:A10 "N/A"
xlex range sort data.xlsx Sheet1 A1:D100 --column B

# Styling
xlex range style data.xlsx Sheet1 A1:D1 --bold --bg-color 4472C4 --text-color FFFFFF
xlex range border data.xlsx Sheet1 A1:D10 --style thin --all
xlex style freeze data.xlsx Sheet1 --rows 1

# Formulas
xlex formula set data.xlsx Sheet1 D1 "SUM(A1:C1)"
xlex formula list data.xlsx Sheet1
xlex formula calc sum data.xlsx Sheet1 A1:A100

# Templates
xlex template apply template.xlsx report.xlsx -D name="Alice" -D date="2026-03-06"

# Convert between formats
xlex convert input.csv output.xlsx
```

## Session Mode

### Batch writes (recommended for automation & AI tools)

Open a file once, make multiple changes, save once. Each command is a separate CLI call — perfect for scripts and AI agents.

```bash
# Start a session
xlex open report.xlsx

# Apply batch commands (each is a separate CLI call)
xlex batch -c "cell set Sheet1 A1 Revenue"
xlex batch -c "cell set Sheet1 B1 Q1"
xlex batch -c "row append Sheet1 Product A,50000,55000"
xlex batch -c "sheet add Summary"

# Check session status
xlex status

# Save all changes back to original file
xlex commit

# Or discard all changes
xlex close
```

Or pipe multiple commands in one shot (fastest):

```bash
xlex batch report.xlsx <<'EOF'
cell set Sheet1 A1 "Revenue"
cell set Sheet1 B1 "Q1"
row append Sheet1 "Product A,50000,55000"
sheet add Summary
cell set Summary A1 "Total"
EOF
```

### Interactive REPL

For large files (>10MB), REPL mode loads the file once for fast repeated reads:

```bash
# Start a REPL session
xlex repl report.xlsx

# In REPL mode:
session> info      # Show workbook information
session> sheets    # List all sheets
session> cell Sheet1 A1        # Get cell value
session> row Sheet1 1          # Get row values
session> search revenue        # Search across all sheets
session> exit      # Exit
```

## Updating

```bash
# Update to the latest version
xlex update

# Check for updates without installing
xlex update --check

# Update to a specific version
xlex update --target v0.3.1
```

## Global Options

| Flag | Short | Effect |
|------|-------|--------|
| `--format` | `-f` | Output: `text` (default), `json`, `csv`, `ndjson` |
| `--dry-run` | | Preview without writing |
| `--output` | `-o` | Write to different file |
| `--quiet` | `-q` | Suppress non-error output |
| `--verbose` | `-v` | Enable verbose output |

## Documentation

For full documentation, visit the [GitHub repository](https://github.com/yen0304/xlex).

## License

MIT
