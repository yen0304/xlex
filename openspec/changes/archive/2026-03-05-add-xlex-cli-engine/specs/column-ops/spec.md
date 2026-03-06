# Column Operations Specification

## ADDED Requirements

### Requirement: Get Column Data

The system SHALL retrieve column data via `xlex column get <file> <sheet> <column>`.

#### Scenario: Get single column
- **GIVEN** the command `xlex column get report.xlsx "Data" A`
- **WHEN** executed
- **THEN** output SHALL be all cell values in column A, one per line

#### Scenario: Get column as JSON
- **GIVEN** the command `xlex column get report.xlsx "Data" A --format json`
- **WHEN** executed
- **THEN** output SHALL be JSON array of cell values

#### Scenario: Get column with row numbers
- **GIVEN** the command `xlex column get report.xlsx "Data" A --with-rows`
- **WHEN** executed
- **THEN** output SHALL include row numbers (e.g., "1:Value1")

#### Scenario: Get column range
- **GIVEN** the command `xlex column get report.xlsx "Data" A:C`
- **WHEN** executed
- **THEN** output SHALL be columns A through C, CSV format

#### Scenario: Get column with limit
- **GIVEN** the command `xlex column get report.xlsx "Data" A --limit 100`
- **WHEN** executed
- **THEN** output SHALL be first 100 values in column A

#### Scenario: Invalid column reference
- **GIVEN** the command `xlex column get report.xlsx "Data" 123`
- **WHEN** executed
- **THEN** error code SHALL be XLEX_E060 (InvalidColumnReference)

### Requirement: Insert Columns

The system SHALL insert columns via `xlex column insert <file> <sheet> <column>`.

#### Scenario: Insert single column
- **GIVEN** the command `xlex column insert report.xlsx "Data" C`
- **WHEN** executed
- **THEN** a new empty column SHALL be inserted at C
- **AND** existing columns C+ SHALL shift right

#### Scenario: Insert multiple columns
- **GIVEN** the command `xlex column insert report.xlsx "Data" C --count 3`
- **WHEN** executed
- **THEN** 3 empty columns SHALL be inserted starting at C

#### Scenario: Formula reference update
- **GIVEN** a formula "=SUM(A1:D1)" in the sheet
- **WHEN** inserting a column at C
- **THEN** the formula SHALL be updated to "=SUM(A1:E1)"

#### Scenario: Insert at end
- **GIVEN** the command `xlex column insert report.xlsx "Data" --at-end`
- **WHEN** executed
- **THEN** a new column SHALL be added after the last used column

### Requirement: Delete Columns

The system SHALL delete columns via `xlex column delete <file> <sheet> <column>`.

#### Scenario: Delete single column
- **GIVEN** the command `xlex column delete report.xlsx "Data" C`
- **WHEN** executed
- **THEN** column C SHALL be deleted
- **AND** columns D+ SHALL shift left

#### Scenario: Delete column range
- **GIVEN** the command `xlex column delete report.xlsx "Data" C:E`
- **WHEN** executed
- **THEN** columns C through E SHALL be deleted

#### Scenario: Delete with formula update
- **GIVEN** a formula "=SUM(A1:E1)" referencing deleted columns
- **WHEN** deleting columns C:D
- **THEN** the formula SHALL be updated to "=SUM(A1:C1)"

#### Scenario: Delete column containing formula reference
- **GIVEN** a formula "=C1" and deleting column C
- **WHEN** executed
- **THEN** the formula SHALL become "=#REF!"

### Requirement: Copy Columns

The system SHALL copy columns via `xlex column copy <file> <sheet> <source> <dest>`.

#### Scenario: Copy single column
- **GIVEN** the command `xlex column copy report.xlsx "Data" A D`
- **WHEN** executed
- **THEN** column A SHALL be copied to column D
- **AND** column D SHALL be inserted (not overwritten)

#### Scenario: Copy column range
- **GIVEN** the command `xlex column copy report.xlsx "Data" A:C F`
- **WHEN** executed
- **THEN** columns A-C SHALL be copied starting at column F

#### Scenario: Copy to another sheet
- **GIVEN** the command `xlex column copy report.xlsx "Data" A --to-sheet "Backup" A`
- **WHEN** executed
- **THEN** column A from Data SHALL be copied to column A of Backup

#### Scenario: Copy with formula adjustment
- **GIVEN** column A contains "=A1+B1"
- **WHEN** copying to column D
- **THEN** the formula SHALL become "=D1+E1"

### Requirement: Move Columns

The system SHALL move columns via `xlex column move <file> <sheet> <source> <dest>`.

#### Scenario: Move single column
- **GIVEN** the command `xlex column move report.xlsx "Data" A D`
- **WHEN** executed
- **THEN** column A SHALL be moved to position D
- **AND** original column A SHALL be removed

#### Scenario: Move column range
- **GIVEN** the command `xlex column move report.xlsx "Data" A:C F`
- **WHEN** executed
- **THEN** columns A-C SHALL be moved to start at column F

#### Scenario: Move left
- **GIVEN** the command `xlex column move report.xlsx "Data" D A`
- **WHEN** executed
- **THEN** column D SHALL be moved to position A

### Requirement: Set Column Width

The system SHALL set column width via `xlex column width <file> <sheet> <column> <width>`.

#### Scenario: Set single column width
- **GIVEN** the command `xlex column width report.xlsx "Data" A 20`
- **WHEN** executed
- **THEN** column A width SHALL be set to 20 characters

#### Scenario: Set column range width
- **GIVEN** the command `xlex column width report.xlsx "Data" A:C 15`
- **WHEN** executed
- **THEN** columns A-C SHALL all have width 15 characters

#### Scenario: Get column width
- **GIVEN** the command `xlex column width report.xlsx "Data" A`
- **WHEN** executed without width value
- **THEN** output SHALL be the current width of column A

#### Scenario: Auto-fit width
- **GIVEN** the command `xlex column width report.xlsx "Data" A --auto`
- **WHEN** executed
- **THEN** column A width SHALL be set to fit content

#### Scenario: Auto-fit all columns
- **GIVEN** the command `xlex column width report.xlsx "Data" --auto-all`
- **WHEN** executed
- **THEN** all used columns SHALL be auto-fitted

### Requirement: Hide Columns

The system SHALL hide columns via `xlex column hide <file> <sheet> <column>`.

#### Scenario: Hide single column
- **GIVEN** the command `xlex column hide report.xlsx "Data" B`
- **WHEN** executed
- **THEN** column B SHALL be hidden

#### Scenario: Hide column range
- **GIVEN** the command `xlex column hide report.xlsx "Data" B:D`
- **WHEN** executed
- **THEN** columns B-D SHALL be hidden

### Requirement: Unhide Columns

The system SHALL unhide columns via `xlex column unhide <file> <sheet> <column>`.

#### Scenario: Unhide single column
- **GIVEN** the command `xlex column unhide report.xlsx "Data" B`
- **WHEN** executed
- **THEN** column B SHALL be visible

#### Scenario: Unhide all columns
- **GIVEN** the command `xlex column unhide report.xlsx "Data" --all`
- **WHEN** executed
- **THEN** all hidden columns SHALL become visible

### Requirement: Rename Column Header

The system SHALL rename column headers via `xlex column header <file> <sheet> <column> <name>`.

#### Scenario: Set column header
- **GIVEN** the command `xlex column header report.xlsx "Data" A "Customer Name"`
- **WHEN** executed
- **THEN** cell A1 SHALL contain "Customer Name"

#### Scenario: Get column header
- **GIVEN** the command `xlex column header report.xlsx "Data" A`
- **WHEN** executed without name
- **THEN** output SHALL be the value in A1

#### Scenario: Set header row
- **GIVEN** the command `xlex column header report.xlsx "Data" A "Name" --row 2`
- **WHEN** executed
- **THEN** cell A2 SHALL contain "Name"

### Requirement: Find Columns

The system SHALL find columns matching criteria via `xlex column find <file> <sheet>`.

#### Scenario: Find by header
- **GIVEN** the command `xlex column find report.xlsx "Data" --header "Email"`
- **WHEN** executed
- **THEN** output SHALL be the column letter(s) with header "Email"

#### Scenario: Find by value
- **GIVEN** the command `xlex column find report.xlsx "Data" --value "Error"`
- **WHEN** executed
- **THEN** output SHALL list column letters containing "Error"

#### Scenario: Find empty columns
- **GIVEN** the command `xlex column find report.xlsx "Data" --empty`
- **WHEN** executed
- **THEN** output SHALL list column letters that are completely empty

### Requirement: Column Statistics

The system SHALL provide column statistics via `xlex column stats <file> <sheet> <column>`.

#### Scenario: Numeric column stats
- **GIVEN** the command `xlex column stats report.xlsx "Data" B`
- **WHEN** column B contains numbers
- **THEN** output SHALL include: count, sum, min, max, average, median

#### Scenario: String column stats
- **GIVEN** the command `xlex column stats report.xlsx "Data" A`
- **WHEN** column A contains strings
- **THEN** output SHALL include: count, unique count, min length, max length

#### Scenario: JSON output
- **GIVEN** the command `xlex column stats report.xlsx "Data" B --format json`
- **WHEN** executed
- **THEN** output SHALL be valid JSON with all statistics
