# xlex

<p align="center">
  <img src="logo.png" alt="xlex logo" width="640">
</p>

<p align="center">
  <strong>A streaming CLI Excel manipulation tool for developers and automation pipelines.</strong>
</p>

<p align="center">
  <a href="https://github.com/yen0304/xlex/actions/workflows/ci.yml"><img src="https://github.com/yen0304/xlex/actions/workflows/ci.yml/badge.svg" alt="CI"></a>
  <a href="https://codecov.io/gh/yen0304/xlex"><img src="https://codecov.io/gh/yen0304/xlex/graph/badge.svg" alt="codecov"></a>
  <a href="https://opensource.org/licenses/MIT"><img src="https://img.shields.io/badge/License-MIT-yellow.svg" alt="License: MIT"></a>
  <a href="https://blog.rust-lang.org/2023/12/28/Rust-1.75.0.html"><img src="https://img.shields.io/badge/MSRV-1.75-blue.svg" alt="MSRV"></a>
</p>

## Features

- **Streaming Architecture**: Handle files up to 200MB without memory exhaustion
- **CLI-First Design**: Composable commands for shell pipelines
- **Multiple Output Formats**: Text, JSON, CSV, NDJSON
- **Template System**: Variable substitution with `{{placeholder}}` syntax
- **Import/Export**: CSV, JSON, YAML, TSV, Markdown support
- **Cross-Platform**: macOS, Linux, Windows binaries

## Installation

### Shell Script (Linux/macOS)

```bash
curl -fsSL https://raw.githubusercontent.com/yen0304/xlex/main/install.sh | bash
```

### npm

```bash
# Global install
npm install -g xlex

# Or use directly with npx (no install needed)
npx xlex info report.xlsx
```

### Cargo

```bash
cargo install xlex-cli
```

### From Source

```bash
git clone https://github.com/yen0304/xlex.git
cd xlex
cargo build --release
# Binary will be at target/release/xlex
```

### Binary Downloads

Download pre-built binaries from the [releases page](https://github.com/yen0304/xlex/releases).

## Quick Start

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

# Template processing
xlex template apply template.xlsx output.xlsx -D name="John" -D date="2026-01-15"

# Interactive mode
xlex interactive
```

## Command Reference

### Workbook Operations

```bash
xlex info <file>              # Display workbook information
xlex validate <file>          # Validate workbook structure
xlex create <file> [sheets]   # Create a new workbook
xlex clone <src> <dest>       # Create a copy
xlex stats <file>             # Display statistics
xlex props <file> [key]       # Get/set properties
```

### Sheet Operations

```bash
xlex sheet list <file>                    # List all sheets
xlex sheet add <file> <name>              # Add a sheet
xlex sheet remove <file> <name>           # Remove a sheet
xlex sheet rename <file> <old> <new>      # Rename a sheet
xlex sheet copy <file> <src> <dest>       # Copy a sheet
xlex sheet hide <file> <name>             # Hide a sheet
xlex sheet unhide <file> <name>           # Unhide a sheet
```

### Cell Operations

```bash
xlex cell get <file> <sheet> <ref>            # Get cell value
xlex cell set <file> <sheet> <ref> <value>    # Set cell value
xlex cell formula <file> <sheet> <ref>        # Get/set formula
xlex cell clear <file> <sheet> <ref>          # Clear cell
xlex cell type <file> <sheet> <ref>           # Get cell type
```

### Row Operations

```bash
xlex row get <file> <sheet> <row>                 # Get row data
xlex row append <file> <sheet> <values...>        # Append a row
xlex row insert <file> <sheet> <row>              # Insert row
xlex row delete <file> <sheet> <row>              # Delete row
xlex row height <file> <sheet> <row> [height]     # Get/set height
```

### Column Operations

```bash
xlex column get <file> <sheet> <col>              # Get column data
xlex column width <file> <sheet> <col> [width]    # Get/set width
xlex column hide <file> <sheet> <col>             # Hide column
xlex column stats <file> <sheet> <col>            # Column statistics
```

### Range Operations

```bash
xlex range get <file> <sheet> <range>             # Get range data
xlex range clear <file> <sheet> <range>           # Clear range
xlex range fill <file> <sheet> <range> <value>    # Fill range
xlex range merge <file> <sheet> <range>           # Merge cells
xlex range validate <file> <sheet> <range> <rule> # Validate data
```

### Import/Export

```bash
# Export
xlex export csv <file> [-s sheet]             # Export to CSV
xlex export json <file> [-s sheet] [--header] # Export to JSON
xlex export markdown <file> [-s sheet]        # Export to Markdown
xlex export yaml <file> [-s sheet]            # Export to YAML
xlex export ndjson <file> [-s sheet]          # Export to NDJSON

# Import
xlex import csv <source> <dest>               # Import CSV
xlex import json <source> <dest>              # Import JSON
xlex import ndjson <source> <dest>            # Import NDJSON

# Convert
xlex convert <source> <dest>                  # Auto-detect formats
```

### Formula Operations

```bash
xlex formula get <file> <sheet> <cell>    # Get formula
xlex formula set <file> <sheet> <cell> <formula>  # Set formula
xlex formula list <file> <sheet>          # List all formulas
xlex formula check <file>                 # Check for errors
xlex formula calc sum <file> <sheet> <range>  # Calculate sum
```

### Template Operations

```bash
xlex template apply <template> <output> -D key=value  # Apply template
xlex template list <template>                         # List placeholders
xlex template validate <template> --vars vars.json    # Validate
```

## Output Formats

Use `-f` or `--format` to specify output format:

```bash
xlex info report.xlsx -f json    # JSON output
xlex info report.xlsx -f csv     # CSV output
xlex info report.xlsx -f text    # Text output (default)
```

## Global Options

```
-q, --quiet        Suppress all output except errors
-v, --verbose      Enable verbose output
-f, --format       Output format (text, json, csv, ndjson)
    --no-color     Disable colored output
    --color        Force colored output
    --json-errors  Output errors as JSON
    --dry-run      Perform a dry run without making changes
-o, --output       Write output to file
```

## Exit Codes

| Code | Description |
|------|-------------|
| 0    | Success |
| 1    | General error |
| 2    | Invalid arguments |
| 3    | File not found |
| 4    | Permission denied |
| 5    | Invalid file format |
| 6    | Sheet not found |
| 7    | Cell reference error |

## Library Usage

```rust
use xlex_core::{Workbook, CellRef, CellValue};

// Open a workbook
let mut workbook = Workbook::open("report.xlsx")?;

// Get a cell value
let value = workbook.get_cell("Sheet1", &CellRef::parse("A1")?)?;
println!("A1: {}", value);

// Set a cell value
workbook.set_cell("Sheet1", CellRef::parse("B1")?, CellValue::Number(42.0))?;

// Save changes
workbook.save()?;
```

## Contributing

See [CONTRIBUTING.md](CONTRIBUTING.md) for guidelines.

## License

MIT License - see [LICENSE](LICENSE) for details.
