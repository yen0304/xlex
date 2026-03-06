# xlex Command Reference

Complete reference for all xlex CLI commands. Organized by domain.

## Table of Contents

- [Workbook](#workbook)
- [Sheet](#sheet)
- [Cell](#cell)
- [Row](#row)
- [Column](#column)
- [Range](#range)
- [Style](#style)
- [Formula](#formula)
- [Template](#template)
- [Import](#import)
- [Export](#export)
- [Utility](#utility)

---

## Workbook

```bash
xlex info     <file>                           # Display workbook info (sheets, properties, size)
xlex validate <file>                           # Validate workbook structure
xlex create   <file> [-s name] [--sheets a,b] [-F]  # Create new workbook (-s: sheet name, --sheets: multiple, -F: overwrite)
xlex clone    <source> <dest> [-F]             # Copy workbook (-F: overwrite if exists)
xlex stats    <file>                           # Row/cell/formula counts
xlex props get <file> [property]               # Get workbook properties (all or specific)
xlex props set <file> <property> <value>       # Set workbook property (title, creator, etc.)
```

## Sheet

```bash
xlex sheet list   <file>                     # List all sheets
xlex sheet add    <file> <name> [-p pos]     # Add sheet (optional position, 0-indexed)
xlex sheet remove <file> <name>              # Remove sheet
xlex sheet rename <file> <old> <new>         # Rename sheet
xlex sheet copy   <file> <source> <dest>     # Duplicate sheet
xlex sheet move   <file> <name> <position>   # Move to position (0-indexed)
xlex sheet hide   <file> <name> [--very]     # Hide (--very = cannot unhide via Excel UI)
xlex sheet unhide <file> <name>              # Unhide
xlex sheet info   <file> <name>              # Sheet details (dimensions, visibility)
xlex sheet active <file> [name]              # Get or set active sheet
```

## Cell

```bash
xlex cell get     <file> <sheet> <ref>                # Get value (e.g., A1)
xlex cell set     <file> <sheet> <ref> <value> [-t type]  # Set value
          # -t: auto (default), string, number, boolean, formula
xlex cell formula  <file> <sheet> <ref> <formula>     # Set formula (without leading =)
xlex cell clear    <file> <sheet> <ref>               # Clear cell
xlex cell type     <file> <sheet> <ref>               # Get cell type
xlex cell batch    <file>                             # Batch ops from stdin (JSON)
```

### Cell comments

```bash
xlex cell comment get    <file> <sheet> <ref>                   # Get comment
xlex cell comment set    <file> <sheet> <ref> <text> [--author]  # Set comment
xlex cell comment remove <file> <sheet> <ref>                   # Remove comment
xlex cell comment list   <file> <sheet>                         # List all comments
```

### Cell hyperlinks

```bash
xlex cell link get    <file> <sheet> <ref>                    # Get hyperlink
xlex cell link set    <file> <sheet> <ref> <url> [--text]     # Set hyperlink (--text for display)
xlex cell link remove <file> <sheet> <ref>                    # Remove hyperlink
```

## Row

Rows are 1-indexed.

```bash
xlex row get     <file> <sheet> <row>                  # Get row data
xlex row append  <file> <sheet> <values>               # Append (comma-separated values)
xlex row insert  <file> <sheet> <row>                  # Insert blank row at position
xlex row delete  <file> <sheet> <row>                  # Delete row
xlex row copy    <file> <sheet> <src_row> <dest_row>   # Copy row
xlex row move    <file> <sheet> <src_row> <dest_row>   # Move row
xlex row height  <file> <sheet> <row> [height]         # Get/set height (points)
xlex row hide    <file> <sheet> <row>                  # Hide row
xlex row unhide  <file> <sheet> <row>                  # Unhide row
xlex row find    <file> <sheet> <pattern> [-c col]     # Find rows matching pattern
```

## Column

Columns use letters: A, B, ..., Z, AA, AB, ...

```bash
xlex column get     <file> <sheet> <col>               # Get column data
xlex column insert  <file> <sheet> <col>               # Insert column
xlex column delete  <file> <sheet> <col>               # Delete column
xlex column copy    <file> <sheet> <src> <dest>        # Copy column
xlex column move    <file> <sheet> <src> <dest>        # Move column
xlex column width   <file> <sheet> <col> [width]       # Get/set width (characters)
xlex column hide    <file> <sheet> <col>               # Hide column
xlex column unhide  <file> <sheet> <col>               # Unhide column
xlex column header  <file> <sheet> <col>               # Get first-row value
xlex column find    <file> <sheet> <pattern>           # Find columns matching pattern
xlex column stats   <file> <sheet> <col>               # Column statistics (min/max/avg/count)
```

## Range

Ranges use A1:B10 notation.

```bash
xlex range get      <file> <sheet> <range>                   # Get range data
xlex range copy     <file> <sheet> <src_range> <dest_cell>   # Copy range to destination
xlex range move     <file> <sheet> <src_range> <dest_cell>   # Move range
xlex range clear    <file> <sheet> <range> [--values-only]   # Clear (optionally keep formatting)
xlex range fill     <file> <sheet> <range> <value>           # Fill all cells with value
xlex range merge    <file> <sheet> <range>                   # Merge cells
xlex range unmerge  <file> <sheet> <range>                   # Unmerge cells
xlex range sort     <file> <sheet> <range> [--column col] [--descending/-d]  # Sort
xlex range filter   <file> <sheet> <range> <column> <value>  # Filter by column value
xlex range validate <file> <sheet> <range> <rule>            # Data validation rule
```

### Range styling

```bash
xlex range style <file> <sheet> <range> [flags]
    --bold --italic --underline              # Font style
    --font <name> --font-size <size>         # Font face/size
    --text-color <hex> --bg-color <hex>      # Colors (e.g., FF0000)
    --align <left|center|right|justify>      # Horizontal alignment
    --valign <top|middle|bottom>             # Vertical alignment
    --wrap                                   # Enable text wrapping
    --number-format <fmt>                    # Custom format (e.g., #,##0.00)
    --percent                                # Format as percentage
    --currency <symbol>                      # Format as currency
    --date-format <fmt>                      # Date format (e.g., YYYY-MM-DD)

xlex range border <file> <sheet> <range> [flags]
    --style <thin|medium|thick|dashed|dotted|double>  # Border style (default: thin)
    --border-color <hex>                               # Border color
    --all --outline --top --bottom --left --right      # Position flags (pick one or more)
    --none                                             # Remove all borders
```

### Named ranges

```bash
xlex range name  <file> <name> <range> [--sheet scope]  # Define named range (global if no --sheet)
xlex range names <file>                                # List all named ranges
```

## Style

```bash
xlex style list    <file>                              # List all style definitions
xlex style get     <file> <id>                         # Get style details by ID
xlex style apply   <file> <sheet> <range> <style_id>   # Apply style by ID
xlex style copy    <file> <sheet> <src_cell> <dest_range>  # Copy style from cell to range
xlex style clear   <file> <sheet> <range>              # Remove style from range
```

### Conditional formatting

```bash
xlex style condition <file> <sheet> <range> [flags]
    --highlight-cells  --gt <val> --lt <val> --eq <val>   # Highlight rules
    --bg-color <hex>                                       # Highlight color
    --color-scale --min <hex> --max <hex>                  # Color scales
    --data-bars --color <hex>                              # Data bars
    --icon-set <name>                                      # Icon sets
    --list / --remove                                      # List or remove rules
```

### Freeze panes

```bash
xlex style freeze <file> <sheet> [--rows <n>] [--cols <n>] [--at <cell>] [--unfreeze]
    # --rows 1 = freeze top row; --cols 1 = freeze first column
    # --at B2 = freeze at specific cell; --unfreeze = remove freeze panes
```

### Presets

```bash
xlex style preset list                                     # List available presets
xlex style preset apply <file> <sheet> <range> <preset>    # Apply preset
```

## Formula

```bash
xlex formula get      <file> <sheet> <cell>              # Get formula from cell
xlex formula set      <file> <sheet> <cell> <formula>    # Set formula (without =)
xlex formula list     <file> <sheet>                     # List all formulas in sheet
xlex formula eval     <file> <sheet> <formula>           # Evaluate formula
xlex formula check    <file> [sheet]                     # Check for formula errors
xlex formula validate <formula>                          # Validate syntax (no file needed)
xlex formula stats    <file> [sheet]                     # Formula statistics
xlex formula refs     <file> <sheet> <cell> [--dependents] [--precedents]
xlex formula replace  <file> <sheet> <find> <replace>    # Replace references in formulas
xlex formula circular <file> [sheet]                     # Detect circular references
```

### Built-in calculations

Quick calculations without writing formulas into cells:

```bash
xlex formula calc sum   <file> <sheet> <range>
xlex formula calc avg   <file> <sheet> <range>
xlex formula calc count <file> <sheet> <range> [--nonempty]
xlex formula calc min   <file> <sheet> <range>
xlex formula calc max   <file> <sheet> <range>
```

## Template

Templates use `{{placeholder}}` syntax.

```bash
xlex template apply    <template> <output> [-D key=value...] [--vars file.json]
    --per-record                          # One output file per record
    --output-pattern "name_{index}.xlsx"  # Filename pattern for per-record
xlex template init     <output> [--template-type report|invoice|data]
xlex template list     <template>                        # List placeholders
xlex template validate <template> [--vars file] [--schema]
xlex template create   <source> <output> [-p cell=name]  # Create template from existing file
xlex template preview  <template> [--vars file] [-D key=value]
```

## Import

```bash
xlex import csv    <source> <dest> [-s sheet] [-d delimiter] [--header]
xlex import json   <source> <dest> [-s sheet]
xlex import tsv    <source> <dest> [-s sheet]
xlex import ndjson <source> <dest> [-s sheet] [--header]
```

## Export

Use `-` as destination to write to stdout.

```bash
xlex export csv      <source> <dest> [-s sheet] [-d delimiter] [--all]
xlex export json     <source> <dest> [-s sheet] [--header] [--all]
xlex export tsv      <source> <dest> [-s sheet] [--all]
xlex export yaml     <source> <dest> [-s sheet] [--all]
xlex export markdown <source> <dest> [-s sheet] [--all]
xlex export ndjson   <source> <dest> [-s sheet] [--header] [--all]
xlex export meta     <source> <dest>
```

## Utility

```bash
xlex convert <input> <output>                  # Auto-detect format by extension
xlex completion <shell>                        # Generate shell completions (bash/zsh/fish/powershell)
xlex config show [--effective]                 # Show configuration
xlex config get <key>                          # Get config value
xlex config set <key> <value>                  # Set config value
xlex config reset                              # Reset configuration to defaults
xlex config init                               # Initialize configuration file
xlex config validate                           # Validate configuration file
xlex batch -f <file> [--continue-on-error]     # Execute batch commands from file
xlex alias list / add <name> <cmd> / remove <name>  # Manage aliases
xlex examples [command] [--all]                # Show usage examples
xlex man [--output-dir <dir>] [--all]          # Generate man pages
xlex version                                   # Version information
xlex interactive                               # Start REPL mode
xlex session <file>                            # Interactive session (file preloaded in memory)
```
