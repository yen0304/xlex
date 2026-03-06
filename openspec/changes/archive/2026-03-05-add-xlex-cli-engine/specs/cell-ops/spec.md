# Cell Operations Specification

## ADDED Requirements

### Requirement: Get Cell Value

The system SHALL retrieve cell values via `xlex cell get <file> <sheet> <ref>`.

#### Scenario: Get string cell
- **GIVEN** the command `xlex cell get report.xlsx "Data" A1`
- **WHEN** A1 contains the string "Hello"
- **THEN** output SHALL be "Hello"

#### Scenario: Get numeric cell
- **GIVEN** the command `xlex cell get report.xlsx "Data" B1`
- **WHEN** B1 contains the number 42.5
- **THEN** output SHALL be "42.5"

#### Scenario: Get formula cell (value)
- **GIVEN** the command `xlex cell get report.xlsx "Data" C1`
- **WHEN** C1 contains formula "=A1+B1" with cached value 100
- **THEN** output SHALL be "100" (the cached value)

#### Scenario: Get formula cell (formula)
- **GIVEN** the command `xlex cell get report.xlsx "Data" C1 --formula`
- **WHEN** C1 contains formula "=A1+B1"
- **THEN** output SHALL be "=A1+B1"

#### Scenario: Get empty cell
- **GIVEN** the command `xlex cell get report.xlsx "Data" Z99`
- **WHEN** Z99 is empty
- **THEN** output SHALL be empty
- **AND** exit code SHALL be 0

#### Scenario: Get cell with type info
- **GIVEN** the command `xlex cell get report.xlsx "Data" A1 --with-type`
- **WHEN** executed
- **THEN** output SHALL include the cell type (string, number, boolean, formula, error)

#### Scenario: JSON output
- **GIVEN** the command `xlex cell get report.xlsx "Data" A1 --format json`
- **WHEN** executed
- **THEN** output SHALL be JSON with fields: ref, value, type, formula (if applicable)

#### Scenario: Invalid cell reference
- **GIVEN** the command `xlex cell get report.xlsx "Data" InvalidRef`
- **WHEN** executed
- **THEN** error code SHALL be XLEX_E020 (InvalidReference)

#### Scenario: Sheet not found
- **GIVEN** the command `xlex cell get report.xlsx "NoSheet" A1`
- **WHEN** executed
- **THEN** error code SHALL be XLEX_E032 (SheetNotFound)

### Requirement: Set Cell Value

The system SHALL set cell values via `xlex cell set <file> <sheet> <ref> <value>`.

#### Scenario: Set string value
- **GIVEN** the command `xlex cell set report.xlsx "Data" A1 "Hello World"`
- **WHEN** executed
- **THEN** A1 SHALL contain the string "Hello World"

#### Scenario: Set numeric value
- **GIVEN** the command `xlex cell set report.xlsx "Data" B1 42.5`
- **WHEN** executed
- **THEN** B1 SHALL contain the number 42.5 (stored as number type)

#### Scenario: Set boolean value
- **GIVEN** the command `xlex cell set report.xlsx "Data" C1 true`
- **WHEN** executed
- **THEN** C1 SHALL contain boolean TRUE

#### Scenario: Set date value
- **GIVEN** the command `xlex cell set report.xlsx "Data" D1 "2024-01-15" --type date`
- **WHEN** executed
- **THEN** D1 SHALL contain the date as Excel serial number
- **AND** appropriate date format SHALL be applied

#### Scenario: Force string type
- **GIVEN** the command `xlex cell set report.xlsx "Data" E1 "123" --type string`
- **WHEN** executed
- **THEN** E1 SHALL contain "123" as a string, not a number

#### Scenario: Set value from stdin
- **GIVEN** the command `echo "Large text" | xlex cell set report.xlsx "Data" A1 --stdin`
- **WHEN** executed
- **THEN** A1 SHALL contain "Large text"

#### Scenario: Output to different file
- **GIVEN** the command `xlex cell set report.xlsx "Data" A1 "New" --output modified.xlsx`
- **WHEN** executed
- **THEN** modified.xlsx SHALL have the new value
- **AND** report.xlsx SHALL remain unchanged

### Requirement: Set Cell Formula

The system SHALL set cell formulas via `xlex cell formula <file> <sheet> <ref> <formula>`.

#### Scenario: Set simple formula
- **GIVEN** the command `xlex cell formula report.xlsx "Data" C1 "=A1+B1"`
- **WHEN** executed
- **THEN** C1 SHALL contain the formula "=A1+B1"

#### Scenario: Set formula with sheet reference
- **GIVEN** the command `xlex cell formula report.xlsx "Summary" A1 "=SUM(Data!B:B)"`
- **WHEN** executed
- **THEN** A1 SHALL contain the cross-sheet formula

#### Scenario: Formula without equals sign
- **GIVEN** the command `xlex cell formula report.xlsx "Data" C1 "A1+B1"`
- **WHEN** executed
- **THEN** the formula SHALL be stored as "=A1+B1" (equals sign auto-added)

#### Scenario: Invalid formula syntax
- **GIVEN** the command `xlex cell formula report.xlsx "Data" C1 "=SUM("`
- **WHEN** executed
- **THEN** error code SHALL be XLEX_E040 (InvalidFormula)
- **AND** message SHALL indicate the syntax error

### Requirement: Clear Cell

The system SHALL clear cell contents via `xlex cell clear <file> <sheet> <ref>`.

#### Scenario: Clear cell value
- **GIVEN** the command `xlex cell clear report.xlsx "Data" A1`
- **WHEN** executed
- **THEN** A1 SHALL be empty (no value, no formula)
- **AND** cell style SHALL be preserved

#### Scenario: Clear cell completely
- **GIVEN** the command `xlex cell clear report.xlsx "Data" A1 --all`
- **WHEN** executed
- **THEN** A1 SHALL be completely cleared (value, formula, and style)

#### Scenario: Clear formula only
- **GIVEN** the command `xlex cell clear report.xlsx "Data" A1 --formula-only`
- **WHEN** A1 has formula with cached value
- **THEN** formula SHALL be removed
- **AND** cached value SHALL be preserved as static value

### Requirement: Get Cell Type

The system SHALL report cell types via `xlex cell type <file> <sheet> <ref>`.

#### Scenario: String type
- **GIVEN** the command `xlex cell type report.xlsx "Data" A1`
- **WHEN** A1 contains a string
- **THEN** output SHALL be "string"

#### Scenario: Number type
- **GIVEN** the command `xlex cell type report.xlsx "Data" B1`
- **WHEN** B1 contains a number
- **THEN** output SHALL be "number"

#### Scenario: Formula type
- **GIVEN** the command `xlex cell type report.xlsx "Data" C1`
- **WHEN** C1 contains a formula
- **THEN** output SHALL be "formula"

#### Scenario: Boolean type
- **GIVEN** the command `xlex cell type report.xlsx "Data" D1`
- **WHEN** D1 contains TRUE or FALSE
- **THEN** output SHALL be "boolean"

#### Scenario: Error type
- **GIVEN** the command `xlex cell type report.xlsx "Data" E1`
- **WHEN** E1 contains #REF! or similar error
- **THEN** output SHALL be "error"

#### Scenario: Empty cell
- **GIVEN** the command `xlex cell type report.xlsx "Data" Z99`
- **WHEN** Z99 is empty
- **THEN** output SHALL be "empty"

### Requirement: Batch Cell Operations

The system SHALL support batch cell operations via `xlex cell batch <file> <sheet>`.

#### Scenario: Batch set from stdin
- **GIVEN** stdin containing:
  ```
  A1,Hello
  B1,42
  C1,=A1&B1
  ```
- **WHEN** executing `xlex cell batch report.xlsx "Data" --set`
- **THEN** all three cells SHALL be set in a single operation

#### Scenario: Batch get to stdout
- **GIVEN** the command `xlex cell batch report.xlsx "Data" --get A1,B1,C1`
- **WHEN** executed
- **THEN** output SHALL be CSV with ref,value pairs

#### Scenario: Batch from JSON
- **GIVEN** stdin containing JSON array of cell operations
- **WHEN** executing `xlex cell batch report.xlsx "Data" --format json`
- **THEN** all operations SHALL be applied

#### Scenario: Batch with errors
- **GIVEN** a batch with some invalid references
- **WHEN** executed with --continue-on-error
- **THEN** valid operations SHALL be applied
- **AND** errors SHALL be reported at the end

### Requirement: Cell Comment Operations

The system SHALL manage cell comments via `xlex cell comment` subcommands.

#### Scenario: Get comment
- **GIVEN** the command `xlex cell comment get report.xlsx "Data" A1`
- **WHEN** A1 has a comment
- **THEN** output SHALL be the comment text

#### Scenario: Set comment
- **GIVEN** the command `xlex cell comment set report.xlsx "Data" A1 "Review needed"`
- **WHEN** executed
- **THEN** A1 SHALL have the comment "Review needed"

#### Scenario: Remove comment
- **GIVEN** the command `xlex cell comment remove report.xlsx "Data" A1`
- **WHEN** executed
- **THEN** the comment on A1 SHALL be removed

#### Scenario: List comments
- **GIVEN** the command `xlex cell comment list report.xlsx "Data"`
- **WHEN** executed
- **THEN** output SHALL list all cells with comments and their text

### Requirement: Cell Hyperlink Operations

The system SHALL manage cell hyperlinks via `xlex cell link` subcommands.

#### Scenario: Get hyperlink
- **GIVEN** the command `xlex cell link get report.xlsx "Data" A1`
- **WHEN** A1 has a hyperlink
- **THEN** output SHALL be the hyperlink URL

#### Scenario: Set hyperlink
- **GIVEN** the command `xlex cell link set report.xlsx "Data" A1 "https://example.com"`
- **WHEN** executed
- **THEN** A1 SHALL have the hyperlink to https://example.com

#### Scenario: Set hyperlink with display text
- **GIVEN** the command `xlex cell link set report.xlsx "Data" A1 "https://example.com" --text "Click here"`
- **WHEN** executed
- **THEN** A1 SHALL display "Click here" and link to https://example.com

#### Scenario: Remove hyperlink
- **GIVEN** the command `xlex cell link remove report.xlsx "Data" A1`
- **WHEN** executed
- **THEN** the hyperlink on A1 SHALL be removed
- **AND** the cell value SHALL be preserved
