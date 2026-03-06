---
name: xlex-agent
description: Manipulate Excel (.xlsx) files via the xlex CLI — read, write, create, style, formula, import/export, template. Use this skill whenever the user mentions Excel files, spreadsheets, .xlsx, CSV-to-Excel conversion, cell values, sheet operations, workbook creation, or any task involving tabular data in Excel format. Also use when the user wants to generate reports, invoices, or data exports as .xlsx files, even if they don't explicitly say "Excel".
---

# xlex Agent Skill

xlex is a streaming CLI tool for Excel manipulation. Every operation is a shell command — no server, no session state. The file path is your handle; read it, modify it, read it again.

Commands follow: `xlex <command> [subcommand] <file> [args...]`

Run `xlex <command> --help` for full argument details. For the complete command reference, see [references/commands.md](references/commands.md).

## Key behaviors

- Write commands modify the file on disk immediately
- Use `--dry-run` to preview changes without writing
- Use `--output other.xlsx` to keep the original intact
- Use `-f json` for structured output — almost always what you want when parsing programmatically
- Rows are 1-indexed numbers; columns are letters (A, B, ..., Z, AA)

## Core workflows

### 1. Explore a file

Always start here. Understand the structure before making changes.

```bash
xlex info report.xlsx -f json            # sheets, properties, file size
xlex sheet list report.xlsx              # sheet names
xlex range get report.xlsx Sheet1 A1:J1 -f json   # header row
xlex range get report.xlsx Sheet1 A1:J5 -f json   # sample rows
```

Why JSON? Text output is for humans. JSON gives you types, nulls, and structure you can reason about.

### 2. Read and write cells

```bash
xlex cell get  data.xlsx Sheet1 A1                    # read
xlex cell set  data.xlsx Sheet1 A1 "Hello"            # write (auto-detect type)
xlex cell set  data.xlsx Sheet1 B1 "42" -t number     # explicit type
xlex cell formula data.xlsx Sheet1 D1 "SUM(A1:C1)"   # formula
xlex cell clear data.xlsx Sheet1 A1                    # clear
```

### 3. Work with ranges

Ranges are your power tool — read blocks, fill, copy, sort, merge in one call.

```bash
xlex range get   data.xlsx Sheet1 A1:D10 -f json
xlex range fill  data.xlsx Sheet1 A1:A10 "N/A"
xlex range copy  data.xlsx Sheet1 A1:C3 E1
xlex range sort  data.xlsx Sheet1 A1:D100 --column B
xlex range merge data.xlsx Sheet1 A1:C1
```

### 4. Rows, columns, sheets

```bash
xlex row append data.xlsx Sheet1 "a,b,c"       # add row at end
xlex row insert data.xlsx Sheet1 3              # insert blank at row 3
xlex column width data.xlsx Sheet1 A 20.0       # set column width
xlex sheet add  data.xlsx NewSheet              # add sheet
xlex sheet rename data.xlsx OldName NewName     # rename
```

### 5. Styling

```bash
xlex range style data.xlsx Sheet1 A1:D1 --bold --bg-color 4472C4 --text-color FFFFFF
xlex range border data.xlsx Sheet1 A1:D10 --style thin --border-color 000000
xlex style freeze data.xlsx Sheet1 --rows 1      # freeze header row
```

### 6. Search across sheets

Global search — like Ctrl+F in Excel. Searches all sheets by default.

```bash
xlex search data.xlsx "revenue"                         # case-insensitive across all sheets
xlex search data.xlsx "error" -s Sheet1                  # restrict to one sheet
xlex search data.xlsx "^\d{4}-" -r                       # regex: find date-like patterns
xlex search data.xlsx "total" -c B                       # only search column B
xlex search data.xlsx "keyword" -n 10 -f json            # first 10 results as JSON
```

### 7. Import / Export

```bash
xlex export csv  data.xlsx output.csv -s Sheet1
xlex export json data.xlsx - -s Sheet1 --header   # stdout, keys from row 1
xlex export markdown data.xlsx - -s Sheet1         # great for showing in chat
xlex import csv  input.csv output.xlsx --header
xlex convert input.csv output.xlsx                 # auto-detect by extension
```

### 8. Templates

`{{placeholder}}` syntax for variable substitution — invoices, reports, batch documents.

```bash
xlex template apply template.xlsx report.xlsx -D name="Alice" -D date="2026-03-06"
xlex template apply template.xlsx out.xlsx --vars records.json --per-record --output-pattern "invoice_{index}.xlsx"
```

## When to reach for references

- **Need exact flags/syntax for a command?** → Read [references/commands.md](references/commands.md)
- **Complex real-world scenario?** → Read [references/examples.md](references/examples.md)
  - Building styled reports, batch processing, pipeline patterns, formula workflows, template generation

## Global options

| Flag | Short | Effect |
|------|-------|--------|
| `--format` | `-f` | Output: `text` (default), `json`, `csv`, `ndjson` |
| `--dry-run` | | Preview without writing |
| `--output` | `-o` | Write to different file |
| `--quiet` | `-q` | Suppress non-error output |
| `--verbose` | `-v` | Enable verbose output |
| `--no-color` | | Disable colored output |
| `--color` | | Force colored output even when piped |
| `--json-errors` | | Errors as JSON for parsing |

## Exit codes

| Code | Meaning |
|------|---------|
| 0 | Success |
| 1 | General error |
| 2 | Invalid arguments |
| 3 | File not found |
| 4 | Permission denied |
| 5 | Invalid file format |
| 6 | Sheet not found |
| 7 | Cell reference error |
