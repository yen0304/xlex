# xlex

<p align="center">
  <img src="logo.png" alt="xlex logo" width="640">
</p>

<p align="center">
  <strong>A CLI Excel tool designed for AI agents — let Copilot, Cursor, Claude and other coding agents read, write, and manipulate Excel files.</strong>
</p>

<p align="center">
  <a href="https://github.com/yen0304/xlex/actions/workflows/ci.yml"><img src="https://github.com/yen0304/xlex/actions/workflows/ci.yml/badge.svg" alt="CI"></a>
  <a href="https://codecov.io/gh/yen0304/xlex"><img src="https://codecov.io/gh/yen0304/xlex/graph/badge.svg" alt="codecov"></a>
  <a href="https://www.npmjs.com/package/xlex"><img src="https://img.shields.io/npm/v/xlex.svg" alt="npm"></a>
  <a href="https://opensource.org/licenses/MIT"><img src="https://img.shields.io/badge/License-MIT-yellow.svg" alt="License: MIT"></a>
  <a href="https://blog.rust-lang.org/2024/07/25/Rust-1.80.0.html"><img src="https://img.shields.io/badge/MSRV-1.80-blue.svg" alt="MSRV"></a>
</p>

<p align="center">
  English | <a href="README.zh-TW.md">繁體中文</a>
</p>

## Why xlex?

AI coding agents (Copilot, Cursor, Claude Code, etc.) can run CLI commands but can't open Excel files directly. xlex bridges this gap — agents use simple CLI commands to read, write, style, and transform `.xlsx` files without any SDK or library integration.

## Features

- **Agent-Friendly**: Structured JSON output, deterministic exit codes, dry-run support
- **Skill Files Included**: Ready-to-use [agent skill files](docs/skills/xlex-agent/) so agents know every command
- **Session Management**: Git-like `open → batch → commit` workflow for multi-step edits
- **Batch Writes**: In-process batch execution — single open/save cycle, ideal for AI agents
- **Streaming Architecture**: Handle files up to 200MB without memory exhaustion
- **Multiple Output Formats**: Text, JSON, CSV, NDJSON
- **Template System**: Variable substitution with `{{placeholder}}` syntax
- **Import/Export**: CSV, JSON, YAML, TSV, Markdown support
- **Cross-Platform**: macOS, Linux, Windows — single binary, zero dependencies

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

## Session Management

xlex provides a **git-like workflow** for multi-step edits: `open → batch → commit`. This is the recommended approach for AI agents — no interactive REPL needed.

```bash
# Open a file for editing (creates a session)
xlex open report.xlsx

# Apply batch commands to the session
xlex batch -c 'cell set Sheet1 A1 "Hello"' -c 'cell set Sheet1 B1 42'

# Check session status
xlex status

# Save changes back to the original file
xlex commit

# Or discard changes
xlex close
```

### Batch Writes

The `batch` command executes multiple write operations in a **single open/save cycle** — no subprocess spawning, no repeated file I/O:

```bash
# Inline commands with -c
xlex batch report.xlsx -c 'cell set Sheet1 A1 "Title"' -c 'row append Sheet1 a b c'

# From a script file
xlex batch report.xlsx -s commands.txt

# Pipe from stdin
echo 'cell set Sheet1 A1 "Hello"' | xlex batch report.xlsx

# With active session (no file argument needed)
xlex open report.xlsx
xlex batch -c 'cell set Sheet1 A1 "Hello"' -c 'sheet add NewSheet'
xlex commit
```

**Supported batch commands:**
- `cell set <sheet> <ref> <value>` — Set cell value
- `cell clear <sheet> <ref>` — Clear cell
- `cell formula <sheet> <ref> <formula>` — Set formula
- `row append <sheet> <values...>` — Append row
- `row insert <sheet> <row>` — Insert empty row
- `row delete <sheet> <row>` — Delete row
- `sheet add <name>` — Add sheet
- `sheet remove <name>` — Remove sheet
- `sheet rename <old> <new>` — Rename sheet

### Interactive REPL

For exploring large files interactively, use the REPL (read-only):

```bash
# Start a REPL session
xlex repl report.xlsx

# In REPL mode:
session> help      # Show available commands
session> info      # Show workbook information
session> sheets    # List all sheets
session> cell Sheet1 A1        # Get cell value
session> cell Sheet1 B2:D5     # Get range values
session> row Sheet1 1          # Get row values
session> exit      # Exit REPL
```

**Benefits:**
- File is loaded only once at session start
- Subsequent commands execute instantly
- Ideal for exploring large workbooks interactively
- Supports JSON output with `--format json`

## AI Agent Integration

xlex ships with **agent skill files** that teach AI coding agents the full command set. Drop them into your project and your agent instantly knows how to manipulate Excel files.

```
docs/skills/xlex-agent/
├── SKILL.md                    # Core overview — start here
└── references/
    ├── commands.md             # Complete CLI command reference
    └── examples.md             # Real-world workflow examples
```

**Compatible agents**: GitHub Copilot, Cursor, Claude Code, Windsurf, or any agent that supports skill/instruction files.

**How to use**: Copy `docs/skills/xlex-agent/` into your project. The agent will automatically discover and follow the skill files when working with Excel tasks.

**What agents can do with xlex**:
- Read/write cells, rows, columns, ranges
- Create workbooks, manage sheets
- Apply styles, formatting, conditional rules
- Import/export CSV, JSON, YAML, Markdown
- Process templates with variable substitution
- Run formulas and calculations

No MCP server, no SDK, no runtime — just CLI commands the agent calls via terminal.

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
xlex sheet move <file> <name> <pos>       # Move sheet to position
xlex sheet hide <file> <name>             # Hide a sheet
xlex sheet unhide <file> <name>           # Unhide a sheet
xlex sheet info <file> <name>             # Show sheet information
xlex sheet active <file> [name]           # Get/set active sheet
```

### Cell Operations

```bash
xlex cell get <file> <sheet> <ref>            # Get cell value
xlex cell set <file> <sheet> <ref> <value>    # Set cell value
xlex cell formula <file> <sheet> <ref> <formula>  # Set formula
xlex cell clear <file> <sheet> <ref>          # Clear cell
xlex cell type <file> <sheet> <ref>           # Get cell type
xlex cell batch <file>                        # Batch operations from stdin
xlex cell comment get <file> <sheet> <ref>    # Get cell comment
xlex cell comment set <file> <sheet> <ref> <text>  # Set comment
xlex cell link get <file> <sheet> <ref>       # Get hyperlink
xlex cell link set <file> <sheet> <ref> <url> # Set hyperlink
```

### Row Operations

```bash
xlex row get <file> <sheet> <row>                 # Get row data
xlex row append <file> <sheet> <values...>        # Append a row
xlex row insert <file> <sheet> <row>              # Insert row
xlex row delete <file> <sheet> <row>              # Delete row
xlex row copy <file> <sheet> <src> <dest>         # Copy row
xlex row move <file> <sheet> <src> <dest>         # Move row
xlex row height <file> <sheet> <row> [height]     # Get/set height
xlex row hide <file> <sheet> <row>                # Hide row
xlex row unhide <file> <sheet> <row>              # Unhide row
xlex row find <file> <sheet> <pattern>            # Find rows
```

### Column Operations

```bash
xlex column get <file> <sheet> <col>              # Get column data
xlex column insert <file> <sheet> <col>           # Insert column
xlex column delete <file> <sheet> <col>           # Delete column
xlex column copy <file> <sheet> <src> <dest>      # Copy column
xlex column move <file> <sheet> <src> <dest>      # Move column
xlex column width <file> <sheet> <col> [width]    # Get/set width
xlex column hide <file> <sheet> <col>             # Hide column
xlex column unhide <file> <sheet> <col>           # Unhide column
xlex column header <file> <sheet> <col>           # Get column header
xlex column find <file> <sheet> <pattern>         # Find columns
xlex column stats <file> <sheet> <col>            # Column statistics
```

### Range Operations

```bash
xlex range get <file> <sheet> <range>             # Get range data
xlex range copy <file> <sheet> <src> <dest>       # Copy range
xlex range move <file> <sheet> <src> <dest>       # Move range
xlex range clear <file> <sheet> <range>           # Clear range
xlex range fill <file> <sheet> <range> <value>    # Fill range
xlex range merge <file> <sheet> <range>           # Merge cells
xlex range unmerge <file> <sheet> <range>         # Unmerge cells
xlex range style <file> <sheet> <range> [opts]    # Apply styling
xlex range border <file> <sheet> <range> [opts]   # Apply borders
xlex range name <file> <name> <range>             # Define named range
xlex range names <file>                           # List named ranges
xlex range validate <file> <sheet> <range> <rule> # Validate data
xlex range sort <file> <sheet> <range> [opts]     # Sort range
```

### Import/Export

```bash
# Export
xlex export csv <file> [-s sheet]             # Export to CSV
xlex export tsv <file> [-s sheet]             # Export to TSV
xlex export json <file> [-s sheet] [--header] # Export to JSON
xlex export markdown <file> [-s sheet]        # Export to Markdown
xlex export yaml <file> [-s sheet]            # Export to YAML
xlex export ndjson <file> [-s sheet]          # Export to NDJSON
xlex export meta <file>                       # Export metadata

# Import
xlex import csv <source> <dest>               # Import CSV
xlex import tsv <source> <dest>               # Import TSV
xlex import json <source> <dest>              # Import JSON
xlex import ndjson <source> <dest>            # Import NDJSON

# Convert
xlex convert <source> <dest>                  # Auto-detect formats
```

### Formula Operations

```bash
xlex formula get <file> <sheet> <cell>            # Get formula
xlex formula set <file> <sheet> <cell> <formula>  # Set formula
xlex formula list <file> <sheet>                  # List all formulas
xlex formula eval <file> <sheet> <formula>        # Evaluate formula
xlex formula check <file>                         # Check for errors
xlex formula validate <formula>                   # Validate syntax
xlex formula stats <file>                         # Formula statistics
xlex formula refs <file> <sheet> <cell>           # Show references
xlex formula replace <file> <sheet> <find> <replace>  # Replace refs
xlex formula circular <file>                      # Detect circular refs
xlex formula calc sum <file> <sheet> <range>      # Calculate sum
xlex formula calc avg <file> <sheet> <range>      # Calculate average
xlex formula calc count <file> <sheet> <range>    # Count values
xlex formula calc min <file> <sheet> <range>      # Get minimum
xlex formula calc max <file> <sheet> <range>      # Get maximum
```

### Template Operations

```bash
xlex template apply <template> <output> -D key=value  # Apply template
xlex template init <output>                           # Create new template
xlex template list <template>                         # List placeholders
xlex template validate <template> --vars vars.json    # Validate
xlex template create <source> <output>                # Create from existing
xlex template preview <template> --vars vars.json     # Preview rendering
```

### Style Operations

```bash
xlex style list <file>                            # List all styles
xlex style get <file> <id>                        # Get style details
xlex style apply <file> <sheet> <range> <id>      # Apply style
xlex style copy <file> <sheet> <src> <dest>       # Copy style
xlex style clear <file> <sheet> <range>           # Clear style
xlex style condition <file> <sheet> <range> [opts]  # Conditional formatting
xlex style freeze <file> <sheet> [opts]           # Freeze panes
xlex style preset list                            # List presets
xlex style preset apply <file> <sheet> <range> <preset>  # Apply preset
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

## Session & Batch Commands

```bash
xlex open <file>                  # Open a workbook for editing (creates a session)
xlex commit                       # Save session changes back to the original file
xlex close                        # Discard session changes and close
xlex status                       # Show current session status
xlex batch [file] -c <cmd>        # Execute inline batch commands
xlex batch [file] -s <script>     # Execute batch commands from script file
xlex repl <file>                  # Start interactive REPL (read-only)
```

## Utility Commands

```bash
xlex completion <shell>           # Generate shell completions (bash, zsh, fish, powershell)
xlex config show                  # Show configuration
xlex config get <key>             # Get config value
xlex config set <key> <value>     # Set config value
xlex alias list                   # List command aliases
xlex alias add <name> <command>   # Add alias
xlex examples [command]           # Show command examples
xlex man                          # Generate man pages
xlex version                      # Display version information
```

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
