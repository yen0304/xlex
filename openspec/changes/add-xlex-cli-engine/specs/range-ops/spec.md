# Range Operations Specification

## ADDED Requirements

### Requirement: Get Range Data

The system SHALL retrieve range data via `xlex range get <file> <sheet> <range>`.

#### Scenario: Get simple range
- **GIVEN** the command `xlex range get report.xlsx "Data" A1:C10`
- **WHEN** executed
- **THEN** output SHALL be CSV format of the range data

#### Scenario: Get range as JSON matrix
- **GIVEN** the command `xlex range get report.xlsx "Data" A1:C10 --format json`
- **WHEN** executed
- **THEN** output SHALL be 2D JSON array

#### Scenario: Get range as JSON records
- **GIVEN** the command `xlex range get report.xlsx "Data" A1:C10 --format json --records`
- **WHEN** row 1 contains headers
- **THEN** output SHALL be JSON array of objects with header keys

#### Scenario: Get range with formulas
- **GIVEN** the command `xlex range get report.xlsx "Data" A1:C10 --formulas`
- **WHEN** executed
- **THEN** output SHALL include formulas instead of values

#### Scenario: Get used range
- **GIVEN** the command `xlex range get report.xlsx "Data" --used`
- **WHEN** executed
- **THEN** output SHALL be the entire used range of the sheet

### Requirement: Copy Range

The system SHALL copy ranges via `xlex range copy <file> <sheet> <source> [dest-sheet] <dest>`.

#### Scenario: Copy within sheet
- **GIVEN** the command `xlex range copy report.xlsx "Data" A1:B10 D1`
- **WHEN** executed
- **THEN** range A1:B10 SHALL be copied starting at D1
- **AND** destination cells SHALL be overwritten

#### Scenario: Copy to another sheet
- **GIVEN** the command `xlex range copy report.xlsx "Data" A1:B10 "Backup" A1`
- **WHEN** executed
- **THEN** range SHALL be copied to Backup sheet at A1

#### Scenario: Copy with formula adjustment
- **GIVEN** source range contains "=A1+B1"
- **WHEN** copying to D1
- **THEN** formula SHALL become "=D1+E1"

#### Scenario: Copy values only
- **GIVEN** the command `xlex range copy report.xlsx "Data" A1:B10 D1 --values-only`
- **WHEN** executed
- **THEN** only values SHALL be copied (no formulas)

#### Scenario: Copy with styles
- **GIVEN** the command `xlex range copy report.xlsx "Data" A1:B10 D1 --with-styles`
- **WHEN** executed
- **THEN** cell styles SHALL also be copied

### Requirement: Move Range

The system SHALL move ranges via `xlex range move <file> <sheet> <source> [dest-sheet] <dest>`.

#### Scenario: Move within sheet
- **GIVEN** the command `xlex range move report.xlsx "Data" A1:B10 D1`
- **WHEN** executed
- **THEN** range A1:B10 SHALL be moved to D1
- **AND** source range SHALL be cleared

#### Scenario: Move to another sheet
- **GIVEN** the command `xlex range move report.xlsx "Data" A1:B10 "Archive" A1`
- **WHEN** executed
- **THEN** range SHALL be moved to Archive sheet
- **AND** source range in Data SHALL be cleared

#### Scenario: Move with reference update
- **GIVEN** a formula "=SUM(A1:B10)" elsewhere
- **WHEN** moving A1:B10 to D1:E10
- **THEN** the formula SHALL be updated to "=SUM(D1:E10)"

### Requirement: Clear Range

The system SHALL clear ranges via `xlex range clear <file> <sheet> <range>`.

#### Scenario: Clear values
- **GIVEN** the command `xlex range clear report.xlsx "Data" A1:D5`
- **WHEN** executed
- **THEN** all values in A1:D5 SHALL be cleared
- **AND** styles SHALL be preserved

#### Scenario: Clear all
- **GIVEN** the command `xlex range clear report.xlsx "Data" A1:D5 --all`
- **WHEN** executed
- **THEN** values, formulas, and styles SHALL be cleared

#### Scenario: Clear formulas only
- **GIVEN** the command `xlex range clear report.xlsx "Data" A1:D5 --formulas-only`
- **WHEN** executed
- **THEN** only formulas SHALL be cleared
- **AND** cached values SHALL be preserved

#### Scenario: Clear styles only
- **GIVEN** the command `xlex range clear report.xlsx "Data" A1:D5 --styles-only`
- **WHEN** executed
- **THEN** only styles SHALL be cleared
- **AND** values and formulas SHALL be preserved

### Requirement: Fill Range

The system SHALL fill ranges via `xlex range fill <file> <sheet> <range>`.

#### Scenario: Fill with value
- **GIVEN** the command `xlex range fill report.xlsx "Data" A1:D5 "N/A"`
- **WHEN** executed
- **THEN** all cells in A1:D5 SHALL contain "N/A"

#### Scenario: Fill down
- **GIVEN** the command `xlex range fill report.xlsx "Data" A1:A10 --down`
- **WHEN** A1 contains a value or formula
- **THEN** A2:A10 SHALL be filled with A1's content (formulas adjusted)

#### Scenario: Fill right
- **GIVEN** the command `xlex range fill report.xlsx "Data" A1:D1 --right`
- **WHEN** A1 contains a value or formula
- **THEN** B1:D1 SHALL be filled with A1's content (formulas adjusted)

#### Scenario: Fill series
- **GIVEN** the command `xlex range fill report.xlsx "Data" A1:A10 --series`
- **WHEN** A1 contains 1
- **THEN** A1:A10 SHALL contain 1, 2, 3, ..., 10

#### Scenario: Fill date series
- **GIVEN** the command `xlex range fill report.xlsx "Data" A1:A10 --series --step day`
- **WHEN** A1 contains a date
- **THEN** A1:A10 SHALL contain consecutive dates

### Requirement: Merge Cells

The system SHALL merge cells via `xlex range merge <file> <sheet> <range>`.

#### Scenario: Merge range
- **GIVEN** the command `xlex range merge report.xlsx "Data" A1:D1`
- **WHEN** executed
- **THEN** cells A1:D1 SHALL be merged
- **AND** content of A1 SHALL be preserved

#### Scenario: Merge with center
- **GIVEN** the command `xlex range merge report.xlsx "Data" A1:D1 --center`
- **WHEN** executed
- **THEN** cells SHALL be merged
- **AND** content SHALL be centered

#### Scenario: Merge non-empty cells
- **GIVEN** the command `xlex range merge report.xlsx "Data" A1:D1`
- **WHEN** multiple cells have values
- **THEN** only A1's value SHALL be preserved
- **AND** a warning SHALL be displayed

#### Scenario: List merged ranges
- **GIVEN** the command `xlex range merge report.xlsx "Data" --list`
- **WHEN** executed
- **THEN** output SHALL list all merged ranges in the sheet

### Requirement: Unmerge Cells

The system SHALL unmerge cells via `xlex range unmerge <file> <sheet> <range>`.

#### Scenario: Unmerge range
- **GIVEN** the command `xlex range unmerge report.xlsx "Data" A1:D1`
- **WHEN** A1:D1 is merged
- **THEN** the merge SHALL be removed
- **AND** value SHALL remain in A1 only

#### Scenario: Unmerge all
- **GIVEN** the command `xlex range unmerge report.xlsx "Data" --all`
- **WHEN** executed
- **THEN** all merged ranges in the sheet SHALL be unmerged

### Requirement: Named Ranges

The system SHALL manage named ranges via `xlex range name` subcommands.

#### Scenario: Create named range
- **GIVEN** the command `xlex range name report.xlsx "SalesData" "Data!A1:D100"`
- **WHEN** executed
- **THEN** a named range "SalesData" SHALL be created

#### Scenario: Get named range
- **GIVEN** the command `xlex range name report.xlsx "SalesData"`
- **WHEN** executed without range
- **THEN** output SHALL be the range reference for "SalesData"

#### Scenario: Delete named range
- **GIVEN** the command `xlex range name report.xlsx "SalesData" --delete`
- **WHEN** executed
- **THEN** the named range SHALL be removed

#### Scenario: Update named range
- **GIVEN** the command `xlex range name report.xlsx "SalesData" "Data!A1:D200"`
- **WHEN** "SalesData" already exists
- **THEN** the range reference SHALL be updated

### Requirement: List Named Ranges

The system SHALL list named ranges via `xlex range names <file>`.

#### Scenario: List all named ranges
- **GIVEN** the command `xlex range names report.xlsx`
- **WHEN** executed
- **THEN** output SHALL list all named ranges with their references

#### Scenario: JSON output
- **GIVEN** the command `xlex range names report.xlsx --format json`
- **WHEN** executed
- **THEN** output SHALL be JSON array of name/range pairs

#### Scenario: Filter by scope
- **GIVEN** the command `xlex range names report.xlsx --scope "Data"`
- **WHEN** executed
- **THEN** output SHALL only include names scoped to "Data" sheet

### Requirement: Range Validation

The system SHALL validate range data via `xlex range validate <file> <sheet> <range>`.

#### Scenario: Validate numeric range
- **GIVEN** the command `xlex range validate report.xlsx "Data" B2:B100 --type number`
- **WHEN** executed
- **THEN** output SHALL list cells that are not numeric

#### Scenario: Validate with regex
- **GIVEN** the command `xlex range validate report.xlsx "Data" A2:A100 --regex "^[A-Z]{2}\d{4}$"`
- **WHEN** executed
- **THEN** output SHALL list cells not matching the pattern

#### Scenario: Validate not empty
- **GIVEN** the command `xlex range validate report.xlsx "Data" A2:D100 --not-empty`
- **WHEN** executed
- **THEN** output SHALL list empty cells in the range

#### Scenario: Validate unique
- **GIVEN** the command `xlex range validate report.xlsx "Data" A2:A100 --unique`
- **WHEN** executed
- **THEN** output SHALL list duplicate values

### Requirement: Range Sort

The system SHALL sort ranges via `xlex range sort <file> <sheet> <range>`.

#### Scenario: Sort by first column
- **GIVEN** the command `xlex range sort report.xlsx "Data" A1:D100`
- **WHEN** executed
- **THEN** rows SHALL be sorted by column A ascending

#### Scenario: Sort by specific column
- **GIVEN** the command `xlex range sort report.xlsx "Data" A1:D100 --by C`
- **WHEN** executed
- **THEN** rows SHALL be sorted by column C

#### Scenario: Sort descending
- **GIVEN** the command `xlex range sort report.xlsx "Data" A1:D100 --desc`
- **WHEN** executed
- **THEN** rows SHALL be sorted descending

#### Scenario: Sort with header
- **GIVEN** the command `xlex range sort report.xlsx "Data" A1:D100 --header`
- **WHEN** executed
- **THEN** row 1 SHALL be preserved as header
- **AND** rows 2+ SHALL be sorted

#### Scenario: Multi-column sort
- **GIVEN** the command `xlex range sort report.xlsx "Data" A1:D100 --by A,C --order asc,desc`
- **WHEN** executed
- **THEN** rows SHALL be sorted by A ascending, then C descending

### Requirement: Range Filter

The system SHALL filter ranges via `xlex range filter <file> <sheet> <range>`.

#### Scenario: Filter by value
- **GIVEN** the command `xlex range filter report.xlsx "Data" A1:D100 --column B --equals "Active"`
- **WHEN** executed
- **THEN** output SHALL be rows where column B equals "Active"

#### Scenario: Filter by condition
- **GIVEN** the command `xlex range filter report.xlsx "Data" A1:D100 --column C --gt 1000`
- **WHEN** executed
- **THEN** output SHALL be rows where column C > 1000

#### Scenario: Filter to new sheet
- **GIVEN** the command `xlex range filter report.xlsx "Data" A1:D100 --column B --equals "Active" --to "Filtered"`
- **WHEN** executed
- **THEN** matching rows SHALL be copied to "Filtered" sheet

#### Scenario: Delete non-matching rows
- **GIVEN** the command `xlex range filter report.xlsx "Data" A1:D100 --column B --equals "Active" --delete-others`
- **WHEN** executed
- **THEN** rows not matching SHALL be deleted from the sheet
