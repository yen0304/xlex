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
```

## Session Mode

For large files (>10MB), session mode loads the file once and keeps it in memory for faster repeated operations:

```bash
# Start a session
xlex session report.xlsx

# In session mode:
session> info      # Show workbook information
session> sheets    # List all sheets
session> cell Sheet1 A1        # Get cell value
session> cell Sheet1 B2:D5     # Get range values
session> row Sheet1 1          # Get row values
session> exit      # Exit session mode
```

**Benefits:**
- File is loaded only once at session start
- Subsequent commands execute instantly
- Ideal for exploring large workbooks interactively

## Documentation

For full documentation, visit the [GitHub repository](https://github.com/yen0304/xlex).

## License

MIT
