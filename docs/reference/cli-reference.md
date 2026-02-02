# CLI Reference

Complete command-line interface reference for xlex.

## Synopsis

```
xlex [OPTIONS] <COMMAND> [ARGS]
```

## Global Options

| Option | Short | Description |
|--------|-------|-------------|
| `--help` | `-h` | Show help message |
| `--version` | `-V` | Show version |
| `--quiet` | `-q` | Suppress output |
| `--verbose` | | Increase verbosity |
| `--format <FORMAT>` | `-f` | Output format (text, json, csv, table) |
| `--no-color` | | Disable colored output |
| `--config <FILE>` | `-c` | Use config file |
| `--json-errors` | | Output errors as JSON |

## Commands

### Workbook Commands

```
xlex info <FILE>                    Show workbook information
xlex validate <FILE>                Validate workbook structure
xlex clone <SRC> <DST>              Clone workbook
xlex create <FILE>                  Create new workbook
xlex props get <FILE> [PROP]        Get document properties
xlex props set <FILE> <PROP> <VAL>  Set document property
xlex stats <FILE>                   Show workbook statistics
```

### Sheet Commands

```
xlex sheet list <FILE>              List all sheets
xlex sheet add <FILE> <NAME>        Add new sheet
xlex sheet remove <FILE> <NAME>     Remove sheet
xlex sheet rename <FILE> <OLD> <NEW> Rename sheet
xlex sheet copy <FILE> <SRC> <DST>  Copy sheet
xlex sheet move <FILE> <NAME> <POS> Move sheet
xlex sheet hide <FILE> <NAME>       Hide sheet
xlex sheet unhide <FILE> <NAME>     Unhide sheet
xlex sheet info <FILE> <NAME>       Sheet details
xlex sheet active <FILE> [NAME]     Get/set active sheet
```

### Cell Commands

```
xlex cell get <FILE> <CELL>         Get cell value
xlex cell set <FILE> <CELL> <VAL>   Set cell value
xlex cell formula get <FILE> <CELL> Get cell formula
xlex cell formula set <FILE> <CELL> <F> Set formula
xlex cell clear <FILE> <CELL>       Clear cell
xlex cell type <FILE> <CELL>        Get cell type
xlex cell batch <FILE>              Batch operations
xlex cell comment get <FILE> <CELL> Get comment
xlex cell comment set <FILE> <CELL> <TEXT> Set comment
xlex cell comment remove <FILE> <CELL> Remove comment
xlex cell comment list <FILE>       List comments
xlex cell link get <FILE> <CELL>    Get hyperlink
xlex cell link set <FILE> <CELL> <URL> Set hyperlink
xlex cell link remove <FILE> <CELL> Remove hyperlink
```

### Row Commands

```
xlex row get <FILE> <ROW>           Get row data
xlex row append <FILE> [VALUES...]  Append row
xlex row insert <FILE> <POS>        Insert row
xlex row delete <FILE> <ROW>        Delete row
xlex row copy <FILE> <SRC> <DST>    Copy row
xlex row move <FILE> <SRC> <DST>    Move row
xlex row height get <FILE> <ROW>    Get row height
xlex row height set <FILE> <ROW> <H> Set row height
xlex row height auto <FILE> <ROW>   Auto-fit height
xlex row hide <FILE> <ROW>          Hide row
xlex row unhide <FILE> <ROW>        Unhide row
xlex row find <FILE> <PATTERN>      Find rows
```

### Column Commands

```
xlex column get <FILE> <COL>        Get column data
xlex column insert <FILE> <COL>     Insert column
xlex column delete <FILE> <COL>     Delete column
xlex column copy <FILE> <SRC> <DST> Copy column
xlex column move <FILE> <SRC> <DST> Move column
xlex column width get <FILE> <COL>  Get width
xlex column width set <FILE> <COL> <W> Set width
xlex column width auto <FILE> <COL> Auto-fit width
xlex column hide <FILE> <COL>       Hide column
xlex column unhide <FILE> <COL>     Unhide column
xlex column header get <FILE> <COL> Get header
xlex column header set <FILE> <COL> <V> Set header
xlex column find <FILE> <PATTERN>   Find columns
xlex column stats <FILE> <COL>      Column statistics
```

### Range Commands

```
xlex range get <FILE> <RANGE>       Get range data
xlex range copy <FILE> <SRC> <DST>  Copy range
xlex range move <FILE> <SRC> <DST>  Move range
xlex range clear <FILE> <RANGE>     Clear range
xlex range fill <FILE> <RANGE> <V>  Fill range
xlex range merge <FILE> <RANGE>     Merge cells
xlex range unmerge <FILE> <RANGE>   Unmerge cells
xlex range name <FILE> <N> <RANGE>  Create named range
xlex range names <FILE>             List named ranges
xlex range validate <FILE> <RANGE>  Add validation
xlex range sort <FILE> <RANGE>      Sort range
xlex range filter <FILE> <RANGE>    Filter range
```

### Style Commands

```
xlex style list <FILE>              List styles
xlex style get <FILE> <CELL>        Get cell style
xlex range style <FILE> <RANGE>     Apply style
xlex range border <FILE> <RANGE>    Apply borders
xlex style preset list              List presets
xlex style preset apply <F> <R> <P> Apply preset
xlex style preset create <NAME>     Create preset
xlex style preset delete <NAME>     Delete preset
xlex style copy <FILE> <SRC> <DST>  Copy style
xlex style clear <FILE> <RANGE>     Clear formatting
xlex style condition <FILE> <RANGE> Conditional format
xlex style freeze <FILE> <CELL>     Freeze panes
```

### Import/Export Commands

```
xlex to csv <FILE>                  Export to CSV
xlex to json <FILE>                 Export to JSON
xlex to ndjson <FILE>               Export to NDJSON
xlex to meta <FILE>                 Export metadata
xlex from csv <CSV> <XLSX>          Import CSV
xlex from json <JSON> <XLSX>        Import JSON
xlex from ndjson <NDJSON> <XLSX>    Import NDJSON
xlex convert <INPUT> <OUTPUT>       Convert formats
```

### Formula Commands

```
xlex formula validate <FORMULA>     Validate formula
xlex formula list <FILE>            List formulas
xlex formula stats <FILE>           Formula statistics
xlex formula refs <FILE> <CELL>     Find dependencies
xlex formula replace <F> <S> <R>    Replace in formulas
```

### Template Commands

```
xlex template apply <TPL> <OUT>     Apply template
xlex template validate <TPL>        Validate template
xlex template init <SRC> <TPL>      Create template
xlex template preview <TPL>         Preview template
```

### Utility Commands

```
xlex completion <SHELL>             Generate completions
xlex config show                    Show configuration
xlex config get <KEY>               Get config value
xlex config set <KEY> <VALUE>       Set config value
xlex config reset [KEY]             Reset config
xlex config init                    Create config file
xlex config validate [FILE]         Validate config
xlex batch <FILE>                   Execute batch file
xlex alias add <NAME> <CMD>         Add alias
xlex alias list                     List aliases
xlex alias remove <NAME>            Remove alias
xlex interactive [FILE]             Interactive mode
xlex man [COMMAND]                  Show manual
```

## Output Formats

### Text (default)

Human-readable output suitable for terminals.

### JSON

Machine-readable JSON output:

```bash
xlex sheet list report.xlsx --format json
```

```json
[
  {"name": "Sheet1", "index": 0, "active": true, "hidden": false},
  {"name": "Data", "index": 1, "active": false, "hidden": false}
]
```

### CSV

Comma-separated values:

```bash
xlex range get report.xlsx A1:D10 --format csv
```

### Table

Formatted ASCII table:

```bash
xlex range get report.xlsx A1:D10 --format table
```

```
┌────────┬────────┬────────┬────────┐
│ Name   │ Age    │ Dept   │ Salary │
├────────┼────────┼────────┼────────┤
│ John   │ 30     │ Sales  │ 50000  │
│ Jane   │ 25     │ IT     │ 60000  │
└────────┴────────┴────────┴────────┘
```

## Exit Codes

| Code | Description |
|------|-------------|
| `0` | Success |
| `1` | General error |
| `2` | Invalid arguments |
| `3` | File not found |
| `4` | Permission denied |
| `5` | Invalid file format |
| `6` | Invalid range/cell reference |
| `7` | Sheet not found |
| `8` | Formula error |
| `9` | Template error |
| `10` | Configuration error |

See [Exit Codes Reference](exit-codes.md) for complete list.

## Environment Variables

| Variable | Description |
|----------|-------------|
| `XLEX_CONFIG` | Configuration file path |
| `XLEX_LOG_FILE` | Log file path |
| `XLEX_LOG_LEVEL` | Logging level |
| `XLEX_NO_COLOR` | Disable colors |
| `XLEX_QUIET` | Suppress output |
| `NO_COLOR` | Standard no-color |

See [Environment Variables Reference](environment-variables.md) for complete list.

## Examples

```bash
# Basic workflow
xlex create report.xlsx --sheets "Data,Summary"
xlex from csv data.csv report.xlsx --sheet "Data"
xlex range style report.xlsx A1:E1 --bold --fill "#4472C4"
xlex to csv report.xlsx --sheet "Data" > export.csv

# Pipeline processing
xlex to json data.xlsx --headers | jq '.[] | select(.status == "active")'

# Batch operations
xlex batch operations.txt --verbose

# Template generation
xlex template apply invoice-template.xlsx output.xlsx --data order.json
```

## See Also

- [Getting Started](../getting-started/installation.md)
- [Configuration File](config-file.md)
- [Error Codes](error-codes.md)
