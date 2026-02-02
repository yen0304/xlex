# Documentation Specification

## ADDED Requirements

### Requirement: Documentation Site Structure

The project SHALL provide comprehensive documentation in the docs/ directory.

#### Scenario: Documentation structure
- **GIVEN** a user accesses documentation
- **WHEN** browsing docs/
- **THEN** the structure SHALL be:
  ```
  docs/
  ├── index.md                    # Home page
  ├── getting-started/
  │   ├── installation.md
  │   ├── quick-start.md
  │   └── first-steps.md
  ├── commands/
  │   ├── index.md               # Command overview
  │   ├── workbook.md
  │   ├── sheet.md
  │   ├── cell.md
  │   ├── row.md
  │   ├── column.md
  │   ├── range.md
  │   ├── style.md
  │   ├── import-export.md
  │   └── utility.md
  ├── guides/
  │   ├── pipelines.md
  │   ├── automation.md
  │   ├── large-files.md
  │   └── error-handling.md
  ├── reference/
  │   ├── cli-reference.md
  │   ├── error-codes.md
  │   ├── exit-codes.md
  │   └── environment-variables.md
  ├── cookbook/
  │   ├── common-tasks.md
  │   ├── data-migration.md
  │   └── reporting.md
  └── development/
      ├── architecture.md
      ├── contributing.md
      └── building.md
  ```

### Requirement: Installation Documentation

The project SHALL document all installation methods in docs/getting-started/installation.md.

#### Scenario: curl installation
- **GIVEN** a user reads installation docs
- **WHEN** looking for curl method
- **THEN** documentation SHALL include:
  ```bash
  curl -fsSL https://xlex.sh/install | sh
  ```
- **AND** explain what the script does
- **AND** show manual verification steps

#### Scenario: Homebrew installation
- **GIVEN** a user reads installation docs
- **WHEN** looking for Homebrew method
- **THEN** documentation SHALL include:
  ```bash
  brew install xlex
  ```
- **AND** mention tap if needed: `brew tap xlex/tap`

#### Scenario: Cargo installation
- **GIVEN** a user reads installation docs
- **WHEN** looking for cargo method
- **THEN** documentation SHALL include:
  ```bash
  cargo install xlex
  ```
- **AND** mention required Rust version

#### Scenario: npx usage
- **GIVEN** a user reads installation docs
- **WHEN** looking for npx method
- **THEN** documentation SHALL include:
  ```bash
  npx xlex sheet list file.xlsx
  ```
- **AND** explain this downloads binary on first run

#### Scenario: Manual download
- **GIVEN** a user reads installation docs
- **WHEN** looking for manual method
- **THEN** documentation SHALL include:
  - Link to GitHub Releases
  - Available platforms
  - How to verify checksums
  - How to add to PATH

#### Scenario: Version verification
- **GIVEN** a user completes installation
- **WHEN** verifying installation
- **THEN** documentation SHALL show:
  ```bash
  xlex --version
  ```

### Requirement: Quick Start Documentation

The project SHALL provide quick start guide in docs/getting-started/quick-start.md.

#### Scenario: Quick start content
- **GIVEN** a new user
- **WHEN** reading quick start
- **THEN** it SHALL cover in under 5 minutes:
  - Installation (one method)
  - View workbook info
  - List sheets
  - Read a cell
  - Update a cell
  - Export to CSV

#### Scenario: Sample file
- **GIVEN** the quick start guide
- **WHEN** following examples
- **THEN** it SHALL either:
  - Provide a downloadable sample.xlsx
  - OR show how to create one with `xlex create`

### Requirement: Workbook Commands Documentation

The project SHALL document all workbook commands in docs/commands/workbook.md.

#### Scenario: xlex info documentation
- **GIVEN** a user reads workbook docs
- **WHEN** looking for info command
- **THEN** documentation SHALL include:
  - Synopsis: `xlex info <file> [options]`
  - Description
  - Options table (--format, --verbose)
  - Examples with output
  - Related commands

#### Scenario: xlex validate documentation
- **GIVEN** a user reads workbook docs
- **WHEN** looking for validate command
- **THEN** documentation SHALL include:
  - Synopsis: `xlex validate <file> [options]`
  - Description
  - Options (--verbose, --strict)
  - Exit codes explanation
  - Examples

#### Scenario: xlex clone documentation
- **GIVEN** a user reads workbook docs
- **WHEN** looking for clone command
- **THEN** documentation SHALL include:
  - Synopsis: `xlex clone <source> <destination> [options]`
  - Options (--force)
  - Examples
  - Error scenarios

#### Scenario: xlex create documentation
- **GIVEN** a user reads workbook docs
- **WHEN** looking for create command
- **THEN** documentation SHALL include:
  - Synopsis: `xlex create <file> [options]`
  - Options (--sheet, --sheets)
  - Examples creating empty workbook
  - Examples with custom sheets

#### Scenario: xlex props documentation
- **GIVEN** a user reads workbook docs
- **WHEN** looking for props command
- **THEN** documentation SHALL include:
  - Synopsis: `xlex props get <file> [property]`
  - Synopsis: `xlex props set <file> <property> <value>`
  - Available properties list
  - Examples

#### Scenario: xlex stats documentation
- **GIVEN** a user reads workbook docs
- **WHEN** looking for stats command
- **THEN** documentation SHALL include:
  - Synopsis: `xlex stats <file> [options]`
  - Output fields explanation
  - JSON output example

### Requirement: Sheet Commands Documentation

The project SHALL document all sheet commands in docs/commands/sheet.md.

#### Scenario: xlex sheet list documentation
- **GIVEN** a user reads sheet docs
- **WHEN** looking for list command
- **THEN** documentation SHALL include:
  - Synopsis: `xlex sheet list <file> [options]`
  - Options (--long, --format)
  - Output format explanation
  - Examples with sample output

#### Scenario: xlex sheet add documentation
- **GIVEN** a user reads sheet docs
- **WHEN** looking for add command
- **THEN** documentation SHALL include:
  - Synopsis: `xlex sheet add <file> <name> [options]`
  - Options (--at)
  - Sheet name restrictions
  - Examples

#### Scenario: xlex sheet remove documentation
- **GIVEN** a user reads sheet docs
- **WHEN** looking for remove command
- **THEN** documentation SHALL include:
  - Synopsis: `xlex sheet remove <file> <name|--index N>`
  - Warning about data loss
  - Examples

#### Scenario: xlex sheet rename documentation
- **GIVEN** a user reads sheet docs
- **WHEN** looking for rename command
- **THEN** documentation SHALL include:
  - Synopsis: `xlex sheet rename <file> <old> <new>`
  - Note about formula reference updates
  - Examples

#### Scenario: xlex sheet copy documentation
- **GIVEN** a user reads sheet docs
- **WHEN** looking for copy command
- **THEN** documentation SHALL include:
  - Synopsis: `xlex sheet copy <file> <source> <dest> [options]`
  - Options (--to)
  - Examples within and across workbooks

#### Scenario: xlex sheet move documentation
- **GIVEN** a user reads sheet docs
- **WHEN** looking for move command
- **THEN** documentation SHALL include:
  - Synopsis: `xlex sheet move <file> <name> <position>`
  - Position explanation (0-indexed, -1 for last)
  - Examples

#### Scenario: xlex sheet hide/unhide documentation
- **GIVEN** a user reads sheet docs
- **WHEN** looking for hide commands
- **THEN** documentation SHALL include:
  - Synopsis for hide and unhide
  - Options (--very, --all)
  - Explanation of hidden vs veryHidden
  - Examples

#### Scenario: xlex sheet info documentation
- **GIVEN** a user reads sheet docs
- **WHEN** looking for info command
- **THEN** documentation SHALL include:
  - Synopsis: `xlex sheet info <file> <name> [options]`
  - Output fields explanation
  - Examples

#### Scenario: xlex sheet active documentation
- **GIVEN** a user reads sheet docs
- **WHEN** looking for active command
- **THEN** documentation SHALL include:
  - Synopsis: `xlex sheet active <file> [name]`
  - Get vs set behavior
  - Examples

### Requirement: Cell Commands Documentation

The project SHALL document all cell commands in docs/commands/cell.md.

#### Scenario: xlex cell get documentation
- **GIVEN** a user reads cell docs
- **WHEN** looking for get command
- **THEN** documentation SHALL include:
  - Synopsis: `xlex cell get <file> <sheet> <ref> [options]`
  - Options (--formula, --with-type, --format)
  - Cell reference format explanation
  - Examples for different cell types

#### Scenario: xlex cell set documentation
- **GIVEN** a user reads cell docs
- **WHEN** looking for set command
- **THEN** documentation SHALL include:
  - Synopsis: `xlex cell set <file> <sheet> <ref> <value> [options]`
  - Options (--type, --stdin, --output)
  - Type inference explanation
  - Examples

#### Scenario: xlex cell formula documentation
- **GIVEN** a user reads cell docs
- **WHEN** looking for formula command
- **THEN** documentation SHALL include:
  - Synopsis: `xlex cell formula <file> <sheet> <ref> <formula>`
  - Formula syntax notes
  - Cross-sheet reference examples
  - Common formula examples

#### Scenario: xlex cell clear documentation
- **GIVEN** a user reads cell docs
- **WHEN** looking for clear command
- **THEN** documentation SHALL include:
  - Synopsis: `xlex cell clear <file> <sheet> <ref> [options]`
  - Options (--all, --formula-only)
  - Examples

#### Scenario: xlex cell type documentation
- **GIVEN** a user reads cell docs
- **WHEN** looking for type command
- **THEN** documentation SHALL include:
  - Synopsis: `xlex cell type <file> <sheet> <ref>`
  - Possible type values
  - Examples

#### Scenario: xlex cell batch documentation
- **GIVEN** a user reads cell docs
- **WHEN** looking for batch command
- **THEN** documentation SHALL include:
  - Synopsis: `xlex cell batch <file> <sheet> [options]`
  - Input format (CSV, JSON)
  - Options (--set, --get, --format, --continue-on-error)
  - Examples with stdin

#### Scenario: xlex cell comment documentation
- **GIVEN** a user reads cell docs
- **WHEN** looking for comment commands
- **THEN** documentation SHALL include:
  - Synopsis for get, set, remove, list
  - Examples

#### Scenario: xlex cell link documentation
- **GIVEN** a user reads cell docs
- **WHEN** looking for link commands
- **THEN** documentation SHALL include:
  - Synopsis for get, set, remove
  - Options (--text)
  - Examples

### Requirement: Row Commands Documentation

The project SHALL document all row commands in docs/commands/row.md.

#### Scenario: xlex row get documentation
- **GIVEN** a user reads row docs
- **WHEN** looking for get command
- **THEN** documentation SHALL include:
  - Synopsis: `xlex row get <file> <sheet> <row|range> [options]`
  - Options (--format, --headers)
  - Examples

#### Scenario: xlex row append documentation
- **GIVEN** a user reads row docs
- **WHEN** looking for append command
- **THEN** documentation SHALL include:
  - Synopsis: `xlex row append <file> <sheet> [options] [-- values...]`
  - Options (--from, --all-strings)
  - Streaming behavior note
  - Examples from stdin, file, and args

#### Scenario: xlex row insert documentation
- **GIVEN** a user reads row docs
- **WHEN** looking for insert command
- **THEN** documentation SHALL include:
  - Synopsis: `xlex row insert <file> <sheet> <position> [options]`
  - Options (--count)
  - Formula reference update note
  - Examples

#### Scenario: xlex row delete documentation
- **GIVEN** a user reads row docs
- **WHEN** looking for delete command
- **THEN** documentation SHALL include:
  - Synopsis: `xlex row delete <file> <sheet> <row|range>`
  - Warning about data loss
  - Formula reference update note
  - Examples

#### Scenario: xlex row copy/move documentation
- **GIVEN** a user reads row docs
- **WHEN** looking for copy/move commands
- **THEN** documentation SHALL include:
  - Synopsis for both commands
  - Options (--to-sheet)
  - Formula adjustment explanation
  - Examples

#### Scenario: xlex row height documentation
- **GIVEN** a user reads row docs
- **WHEN** looking for height command
- **THEN** documentation SHALL include:
  - Synopsis: `xlex row height <file> <sheet> <row|range> [height]`
  - Options (--auto)
  - Get vs set behavior
  - Examples

#### Scenario: xlex row hide/unhide documentation
- **GIVEN** a user reads row docs
- **WHEN** looking for hide commands
- **THEN** documentation SHALL include:
  - Synopsis for hide and unhide
  - Options (--all)
  - Examples

#### Scenario: xlex row find documentation
- **GIVEN** a user reads row docs
- **WHEN** looking for find command
- **THEN** documentation SHALL include:
  - Synopsis: `xlex row find <file> <sheet> [options]`
  - Options (--value, --column, --regex, --empty)
  - Examples

### Requirement: Column Commands Documentation

The project SHALL document all column commands in docs/commands/column.md.

#### Scenario: xlex column get documentation
- **GIVEN** a user reads column docs
- **WHEN** looking for get command
- **THEN** documentation SHALL include:
  - Synopsis: `xlex column get <file> <sheet> <column|range> [options]`
  - Options (--format, --with-rows, --limit)
  - Column reference format (A, AA, etc.)
  - Examples

#### Scenario: xlex column insert documentation
- **GIVEN** a user reads column docs
- **WHEN** looking for insert command
- **THEN** documentation SHALL include:
  - Synopsis: `xlex column insert <file> <sheet> <column> [options]`
  - Options (--count, --at-end)
  - Formula reference update note
  - Examples

#### Scenario: xlex column delete documentation
- **GIVEN** a user reads column docs
- **WHEN** looking for delete command
- **THEN** documentation SHALL include:
  - Synopsis: `xlex column delete <file> <sheet> <column|range>`
  - Warning about data loss
  - Examples

#### Scenario: xlex column copy/move documentation
- **GIVEN** a user reads column docs
- **WHEN** looking for copy/move commands
- **THEN** documentation SHALL include:
  - Synopsis for both commands
  - Options (--to-sheet)
  - Formula adjustment explanation
  - Examples

#### Scenario: xlex column width documentation
- **GIVEN** a user reads column docs
- **WHEN** looking for width command
- **THEN** documentation SHALL include:
  - Synopsis: `xlex column width <file> <sheet> <column|range> [width]`
  - Options (--auto, --auto-all)
  - Width unit explanation
  - Examples

#### Scenario: xlex column hide/unhide documentation
- **GIVEN** a user reads column docs
- **WHEN** looking for hide commands
- **THEN** documentation SHALL include:
  - Synopsis for hide and unhide
  - Options (--all)
  - Examples

#### Scenario: xlex column header documentation
- **GIVEN** a user reads column docs
- **WHEN** looking for header command
- **THEN** documentation SHALL include:
  - Synopsis: `xlex column header <file> <sheet> <column> [name]`
  - Options (--row)
  - Get vs set behavior
  - Examples

#### Scenario: xlex column find documentation
- **GIVEN** a user reads column docs
- **WHEN** looking for find command
- **THEN** documentation SHALL include:
  - Synopsis: `xlex column find <file> <sheet> [options]`
  - Options (--header, --value, --empty)
  - Examples

#### Scenario: xlex column stats documentation
- **GIVEN** a user reads column docs
- **WHEN** looking for stats command
- **THEN** documentation SHALL include:
  - Synopsis: `xlex column stats <file> <sheet> <column> [options]`
  - Output fields for numeric vs string columns
  - Examples

### Requirement: Range Commands Documentation

The project SHALL document all range commands in docs/commands/range.md.

#### Scenario: xlex range get documentation
- **GIVEN** a user reads range docs
- **WHEN** looking for get command
- **THEN** documentation SHALL include:
  - Synopsis: `xlex range get <file> <sheet> <range> [options]`
  - Options (--format, --records, --formulas, --used)
  - Range notation explanation
  - Examples

#### Scenario: xlex range copy documentation
- **GIVEN** a user reads range docs
- **WHEN** looking for copy command
- **THEN** documentation SHALL include:
  - Synopsis: `xlex range copy <file> <sheet> <source> [dest-sheet] <dest>`
  - Options (--values-only, --with-styles)
  - Formula adjustment explanation
  - Examples

#### Scenario: xlex range move documentation
- **GIVEN** a user reads range docs
- **WHEN** looking for move command
- **THEN** documentation SHALL include:
  - Synopsis: `xlex range move <file> <sheet> <source> [dest-sheet] <dest>`
  - Reference update explanation
  - Examples

#### Scenario: xlex range clear documentation
- **GIVEN** a user reads range docs
- **WHEN** looking for clear command
- **THEN** documentation SHALL include:
  - Synopsis: `xlex range clear <file> <sheet> <range> [options]`
  - Options (--all, --formulas-only, --styles-only)
  - Examples

#### Scenario: xlex range fill documentation
- **GIVEN** a user reads range docs
- **WHEN** looking for fill command
- **THEN** documentation SHALL include:
  - Synopsis: `xlex range fill <file> <sheet> <range> [value] [options]`
  - Options (--down, --right, --series, --step)
  - Series fill explanation
  - Examples

#### Scenario: xlex range merge/unmerge documentation
- **GIVEN** a user reads range docs
- **WHEN** looking for merge commands
- **THEN** documentation SHALL include:
  - Synopsis for merge and unmerge
  - Options (--center, --list, --all)
  - Data loss warning for merge
  - Examples

#### Scenario: xlex range name documentation
- **GIVEN** a user reads range docs
- **WHEN** looking for name command
- **THEN** documentation SHALL include:
  - Synopsis: `xlex range name <file> <name> [range] [options]`
  - Options (--delete)
  - Get, set, delete behavior
  - Examples

#### Scenario: xlex range names documentation
- **GIVEN** a user reads range docs
- **WHEN** looking for names command
- **THEN** documentation SHALL include:
  - Synopsis: `xlex range names <file> [options]`
  - Options (--format, --scope)
  - Examples

#### Scenario: xlex range validate documentation
- **GIVEN** a user reads range docs
- **WHEN** looking for validate command
- **THEN** documentation SHALL include:
  - Synopsis: `xlex range validate <file> <sheet> <range> [options]`
  - Options (--type, --regex, --not-empty, --unique)
  - Examples

#### Scenario: xlex range sort documentation
- **GIVEN** a user reads range docs
- **WHEN** looking for sort command
- **THEN** documentation SHALL include:
  - Synopsis: `xlex range sort <file> <sheet> <range> [options]`
  - Options (--by, --desc, --header, --order)
  - Multi-column sort explanation
  - Examples

#### Scenario: xlex range filter documentation
- **GIVEN** a user reads range docs
- **WHEN** looking for filter command
- **THEN** documentation SHALL include:
  - Synopsis: `xlex range filter <file> <sheet> <range> [options]`
  - Options (--column, --equals, --gt, --lt, --to, --delete-others)
  - Examples

### Requirement: Style Commands Documentation

The project SHALL document all style commands in docs/commands/style.md.

#### Scenario: xlex style list documentation
- **GIVEN** a user reads style docs
- **WHEN** looking for list command
- **THEN** documentation SHALL include:
  - Synopsis: `xlex style list <file> [options]`
  - Options (--long, --format)
  - Examples

#### Scenario: xlex style get documentation
- **GIVEN** a user reads style docs
- **WHEN** looking for get command
- **THEN** documentation SHALL include:
  - Synopsis: `xlex style get <file> <sheet> <ref> [options]`
  - Options (--id-only, --format)
  - Output fields explanation
  - Examples

#### Scenario: xlex range style documentation
- **GIVEN** a user reads style docs
- **WHEN** looking for range style command
- **THEN** documentation SHALL include:
  - Synopsis: `xlex range style <file> <sheet> <range> [options]`
  - All style options with descriptions:
    - --bold, --italic, --underline
    - --font, --font-size, --color
    - --bg-color, --align, --valign
    - --wrap, --number-format, --date-format
    - --percent, --currency
  - Color format explanation (#RRGGBB)
  - Examples

#### Scenario: xlex range border documentation
- **GIVEN** a user reads style docs
- **WHEN** looking for border command
- **THEN** documentation SHALL include:
  - Synopsis: `xlex range border <file> <sheet> <range> [options]`
  - Options (--all, --outline, --top, --bottom, --left, --right, --style, --border-color, --none)
  - Border styles list
  - Examples

#### Scenario: xlex style preset documentation
- **GIVEN** a user reads style docs
- **WHEN** looking for preset commands
- **THEN** documentation SHALL include:
  - Synopsis for list, apply, create, delete
  - Built-in presets list
  - Custom preset creation examples

#### Scenario: xlex style copy documentation
- **GIVEN** a user reads style docs
- **WHEN** looking for copy command
- **THEN** documentation SHALL include:
  - Synopsis: `xlex style copy <file> <sheet> <source> <dest> [options]`
  - Options (--to-sheet)
  - Examples

#### Scenario: xlex style clear documentation
- **GIVEN** a user reads style docs
- **WHEN** looking for clear command
- **THEN** documentation SHALL include:
  - Synopsis: `xlex style clear <file> <sheet> <range> [options]`
  - Options (--font-only)
  - Examples

#### Scenario: xlex style condition documentation
- **GIVEN** a user reads style docs
- **WHEN** looking for condition command
- **THEN** documentation SHALL include:
  - Synopsis: `xlex style condition <file> <sheet> <range> [options]`
  - Options (--highlight-cells, --gt, --lt, --eq, --bg-color, --color-scale, --data-bars, --icon-set, --list, --remove)
  - Conditional formatting types explanation
  - Examples

#### Scenario: xlex style freeze documentation
- **GIVEN** a user reads style docs
- **WHEN** looking for freeze command
- **THEN** documentation SHALL include:
  - Synopsis: `xlex style freeze <file> <sheet> [options]`
  - Options (--rows, --cols, --at, --unfreeze)
  - Examples

### Requirement: Import/Export Commands Documentation

The project SHALL document all import/export commands in docs/commands/import-export.md.

#### Scenario: xlex to csv documentation
- **GIVEN** a user reads import/export docs
- **WHEN** looking for CSV export
- **THEN** documentation SHALL include:
  - Synopsis: `xlex to csv <file> <sheet> [options]`
  - Options (--output, --range, --delimiter, --quote, --no-header, --formulas)
  - Streaming behavior note
  - Examples

#### Scenario: xlex to json documentation
- **GIVEN** a user reads import/export docs
- **WHEN** looking for JSON export
- **THEN** documentation SHALL include:
  - Synopsis: `xlex to json <file> <sheet> [options]`
  - Options (--output, --records, --pretty, --range, --with-metadata, --null-value)
  - Output format examples (array vs records)
  - Examples

#### Scenario: xlex to ndjson documentation
- **GIVEN** a user reads import/export docs
- **WHEN** looking for NDJSON export
- **THEN** documentation SHALL include:
  - Synopsis: `xlex to ndjson <file> <sheet> [options]`
  - Options (--headers)
  - Streaming benefits explanation
  - Examples

#### Scenario: xlex from csv documentation
- **GIVEN** a user reads import/export docs
- **WHEN** looking for CSV import
- **THEN** documentation SHALL include:
  - Synopsis: `xlex from csv <file> <sheet> [options]`
  - Options (--input, --append, --replace, --delimiter, --all-strings, --header)
  - Type inference explanation
  - Examples from stdin and file

#### Scenario: xlex from json documentation
- **GIVEN** a user reads import/export docs
- **WHEN** looking for JSON import
- **THEN** documentation SHALL include:
  - Synopsis: `xlex from json <file> <sheet> [options]`
  - Options (--input, --path, --flatten)
  - Input format explanation (array of arrays vs objects)
  - Examples

#### Scenario: xlex from ndjson documentation
- **GIVEN** a user reads import/export docs
- **WHEN** looking for NDJSON import
- **THEN** documentation SHALL include:
  - Synopsis: `xlex from ndjson <file> <sheet> [options]`
  - Streaming benefits explanation
  - Examples

#### Scenario: xlex convert documentation
- **GIVEN** a user reads import/export docs
- **WHEN** looking for convert command
- **THEN** documentation SHALL include:
  - Synopsis: `xlex convert <input> <output> [options]`
  - Supported format pairs
  - Options (--all)
  - Examples

#### Scenario: xlex to meta documentation
- **GIVEN** a user reads import/export docs
- **WHEN** looking for metadata export
- **THEN** documentation SHALL include:
  - Synopsis: `xlex to meta <file> [options]`
  - Output fields explanation
  - Examples

### Requirement: Utility Commands Documentation

The project SHALL document utility commands in docs/commands/utility.md.

#### Scenario: xlex completion documentation
- **GIVEN** a user reads utility docs
- **WHEN** looking for completion command
- **THEN** documentation SHALL include:
  - Synopsis: `xlex completion <shell>`
  - Supported shells (bash, zsh, fish, powershell)
  - Installation instructions for each shell
  - Examples

#### Scenario: xlex config documentation
- **GIVEN** a user reads utility docs
- **WHEN** looking for config command
- **THEN** documentation SHALL include:
  - Synopsis for show, get, set, reset
  - Available configuration options
  - Config file location
  - Examples

#### Scenario: xlex batch documentation
- **GIVEN** a user reads utility docs
- **WHEN** looking for batch command
- **THEN** documentation SHALL include:
  - Synopsis: `xlex batch [options]`
  - Options (--file, --continue-on-error, --transaction)
  - Input format explanation
  - Examples

#### Scenario: xlex alias documentation
- **GIVEN** a user reads utility docs
- **WHEN** looking for alias command
- **THEN** documentation SHALL include:
  - Synopsis for add, list, remove
  - Built-in aliases list
  - Examples

#### Scenario: xlex interactive documentation
- **GIVEN** a user reads utility docs
- **WHEN** looking for interactive command
- **THEN** documentation SHALL include:
  - Synopsis: `xlex interactive <file>`
  - Available commands in interactive mode
  - Keyboard shortcuts
  - Examples

### Requirement: Error Codes Reference

The project SHALL document all error codes in docs/reference/error-codes.md.

#### Scenario: Error codes table
- **GIVEN** a user reads error codes reference
- **WHEN** looking up an error
- **THEN** documentation SHALL include a table with:
  - Error code (XLEX_EXXX)
  - Category
  - Description
  - Common causes
  - Resolution suggestions

#### Scenario: Error categories
- **GIVEN** a user reads error codes reference
- **WHEN** browsing errors
- **THEN** errors SHALL be grouped by category:
  - File errors (E001-E010)
  - Parse errors (E011-E019)
  - Reference errors (E020-E029)
  - Sheet errors (E030-E039)
  - Formula errors (E040-E049)
  - Row errors (E050-E059)
  - Column errors (E060-E069)
  - Import/Export errors (E070-E079)
  - Style errors (E080-E089)

### Requirement: Environment Variables Reference

The project SHALL document environment variables in docs/reference/environment-variables.md.

#### Scenario: Environment variables table
- **GIVEN** a user reads env vars reference
- **WHEN** looking for configuration options
- **THEN** documentation SHALL include:
  - Variable name
  - Description
  - Default value
  - Example values

#### Scenario: Environment variables list
- **GIVEN** a user reads env vars reference
- **WHEN** browsing variables
- **THEN** documentation SHALL cover:
  - XLEX_DEFAULT_FORMAT
  - XLEX_NO_COLOR
  - XLEX_QUIET
  - XLEX_STRING_CACHE_SIZE
  - XLEX_CONFIG
  - XLEX_LOG_FILE

### Requirement: Pipeline Guide

The project SHALL provide pipeline integration guide in docs/guides/pipelines.md.

#### Scenario: Pipeline guide content
- **GIVEN** a user reads pipeline guide
- **WHEN** learning about integration
- **THEN** documentation SHALL include:
  - Unix philosophy explanation
  - stdin/stdout patterns
  - Examples with common tools (jq, awk, grep)
  - Database integration examples
  - CI/CD integration examples

#### Scenario: Database examples
- **GIVEN** a user reads pipeline guide
- **WHEN** looking for database integration
- **THEN** documentation SHALL include examples for:
  - PostgreSQL (COPY TO/FROM)
  - MySQL
  - SQLite
  - MongoDB (via JSON)

### Requirement: Cookbook

The project SHALL provide a cookbook in docs/cookbook/.

#### Scenario: Common tasks cookbook
- **GIVEN** a user reads cookbook
- **WHEN** looking for common tasks
- **THEN** documentation SHALL include recipes for:
  - Create report from template
  - Merge multiple xlsx files
  - Split xlsx by sheet
  - Find and replace across sheets
  - Generate summary sheet
  - Apply consistent formatting
  - Validate data before import

#### Scenario: Data migration cookbook
- **GIVEN** a user reads cookbook
- **WHEN** looking for migration tasks
- **THEN** documentation SHALL include recipes for:
  - CSV to xlsx conversion
  - xlsx to database import
  - Database to xlsx export
  - JSON API to xlsx
  - xlsx to JSON API

### Requirement: Documentation Site Generation

The project SHALL support documentation site generation.

#### Scenario: MkDocs configuration
- **GIVEN** the docs/ directory
- **WHEN** building documentation site
- **THEN** mkdocs.yml SHALL be provided with:
  - Site name and description
  - Theme configuration (Material for MkDocs)
  - Navigation structure
  - Search configuration
  - Code highlighting

#### Scenario: Documentation hosting
- **GIVEN** documentation is built
- **WHEN** deploying
- **THEN** documentation SHALL be deployable to:
  - GitHub Pages
  - Read the Docs
  - Custom domain (docs.xlex.sh)
