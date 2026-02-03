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

## Documentation

For full documentation, visit the [GitHub repository](https://github.com/yen0304/xlex).

## License

MIT
