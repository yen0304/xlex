# Formula Operations Specification

## ADDED Requirements

### Requirement: Validate Formula Syntax

The system SHALL validate formula syntax via `xlex formula validate <file>`.

#### Scenario: Validate all formulas
- **GIVEN** the command `xlex formula validate report.xlsx`
- **WHEN** executed
- **THEN** all formulas in the workbook SHALL be syntax-checked
- **AND** output SHALL list any invalid formulas with locations

#### Scenario: Validate specific sheet
- **GIVEN** the command `xlex formula validate report.xlsx --sheet "Data"`
- **WHEN** executed
- **THEN** only formulas in "Data" sheet SHALL be validated

#### Scenario: Valid formulas
- **GIVEN** a workbook with all valid formulas
- **WHEN** validated
- **THEN** exit code SHALL be 0
- **AND** output SHALL indicate "All formulas valid"

#### Scenario: Invalid formula detected
- **GIVEN** a workbook with formula "=SUM(A1:A10" (missing parenthesis)
- **WHEN** validated
- **THEN** exit code SHALL be non-zero
- **AND** output SHALL show:
  - Cell reference (e.g., Sheet1!B5)
  - Invalid formula text
  - Error description

#### Scenario: JSON output
- **GIVEN** the command `xlex formula validate report.xlsx --format json`
- **WHEN** executed
- **THEN** output SHALL be JSON with validation results

#### Scenario: Check circular references
- **GIVEN** the command `xlex formula validate report.xlsx --check-circular`
- **WHEN** executed
- **THEN** potential circular references SHALL be detected and reported

#### Scenario: Strict mode
- **GIVEN** the command `xlex formula validate report.xlsx --strict`
- **WHEN** executed
- **THEN** additional checks SHALL be performed:
  - Unknown function names
  - References to non-existent sheets
  - References outside data range

### Requirement: List Formulas

The system SHALL list all formulas via `xlex formula list <file>`.

#### Scenario: List all formulas
- **GIVEN** the command `xlex formula list report.xlsx`
- **WHEN** executed
- **THEN** output SHALL list all formulas with:
  - Cell reference
  - Formula text
  - Cached value (if available)

#### Scenario: List by sheet
- **GIVEN** the command `xlex formula list report.xlsx --sheet "Data"`
- **WHEN** executed
- **THEN** only formulas in "Data" sheet SHALL be listed

#### Scenario: JSON output
- **GIVEN** the command `xlex formula list report.xlsx --format json`
- **WHEN** executed
- **THEN** output SHALL be JSON array of formula objects

#### Scenario: Count only
- **GIVEN** the command `xlex formula list report.xlsx --count`
- **WHEN** executed
- **THEN** output SHALL be only the count of formulas

#### Scenario: Filter by function
- **GIVEN** the command `xlex formula list report.xlsx --function SUM`
- **WHEN** executed
- **THEN** only formulas containing SUM function SHALL be listed

#### Scenario: Show dependencies
- **GIVEN** the command `xlex formula list report.xlsx --with-deps`
- **WHEN** executed
- **THEN** output SHALL include cells that each formula depends on

#### Scenario: No formulas found
- **GIVEN** a workbook with no formulas
- **WHEN** listing formulas
- **THEN** output SHALL indicate "No formulas found"
- **AND** exit code SHALL be 0

### Requirement: Formula Statistics

The system SHALL provide formula statistics via `xlex formula stats <file>`.

#### Scenario: Show statistics
- **GIVEN** the command `xlex formula stats report.xlsx`
- **WHEN** executed
- **THEN** output SHALL include:
  - Total formula count
  - Per-sheet breakdown
  - Most used functions
  - Deepest dependency chain
  - Cross-sheet reference count

#### Scenario: JSON output
- **GIVEN** the command `xlex formula stats report.xlsx --format json`
- **WHEN** executed
- **THEN** output SHALL be valid JSON with all statistics

### Requirement: Find Formula References

The system SHALL find cells referencing a target via `xlex formula refs <file> <ref>`.

#### Scenario: Find dependents
- **GIVEN** the command `xlex formula refs report.xlsx "Data!A1"`
- **WHEN** executed
- **THEN** output SHALL list all cells with formulas referencing Data!A1

#### Scenario: Find precedents
- **GIVEN** the command `xlex formula refs report.xlsx "Data!B1" --precedents`
- **WHEN** executed
- **THEN** output SHALL list all cells that B1's formula depends on

#### Scenario: Recursive dependencies
- **GIVEN** the command `xlex formula refs report.xlsx "Data!A1" --recursive`
- **WHEN** executed
- **THEN** output SHALL show full dependency tree

### Requirement: Replace Formula References

The system SHALL replace references in formulas via `xlex formula replace <file>`.

#### Scenario: Replace sheet reference
- **GIVEN** the command `xlex formula replace report.xlsx --from "OldSheet" --to "NewSheet"`
- **WHEN** executed
- **THEN** all formulas referencing "OldSheet" SHALL be updated to "NewSheet"

#### Scenario: Replace cell reference
- **GIVEN** the command `xlex formula replace report.xlsx --from "A1:A100" --to "B1:B100"`
- **WHEN** executed
- **THEN** all formulas referencing A1:A100 SHALL be updated

#### Scenario: Dry run
- **GIVEN** the command `xlex formula replace report.xlsx --from "X" --to "Y" --dry-run`
- **WHEN** executed
- **THEN** changes SHALL be shown but not applied

#### Scenario: Preview changes
- **GIVEN** the command with --dry-run
- **WHEN** executed
- **THEN** output SHALL show before/after for each affected formula
