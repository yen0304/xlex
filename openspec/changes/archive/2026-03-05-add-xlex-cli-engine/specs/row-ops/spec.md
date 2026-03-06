# Row Operations Specification

## ADDED Requirements

### Requirement: Get Row Data

The system SHALL retrieve row data via `xlex row get <file> <sheet> <row>`.

#### Scenario: Get single row
- **GIVEN** the command `xlex row get report.xlsx "Data" 1`
- **WHEN** executed
- **THEN** output SHALL be all cell values in row 1, comma-separated

#### Scenario: Get row as JSON
- **GIVEN** the command `xlex row get report.xlsx "Data" 1 --format json`
- **WHEN** executed
- **THEN** output SHALL be JSON array of cell values

#### Scenario: Get row with headers
- **GIVEN** the command `xlex row get report.xlsx "Data" 2 --headers`
- **WHEN** row 1 contains headers
- **THEN** output SHALL be JSON object with header keys

#### Scenario: Get row range
- **GIVEN** the command `xlex row get report.xlsx "Data" 1:10`
- **WHEN** executed
- **THEN** output SHALL be rows 1 through 10, one per line

#### Scenario: Row out of bounds
- **GIVEN** the command `xlex row get report.xlsx "Data" 0`
- **WHEN** executed (rows are 1-indexed)
- **THEN** error code SHALL be XLEX_E050 (InvalidRowIndex)

### Requirement: Append Rows

The system SHALL append rows via `xlex row append <file> <sheet>`.

#### Scenario: Append single row from args
- **GIVEN** the command `xlex row append report.xlsx "Data" -- "Value1" "Value2" "Value3"`
- **WHEN** executed
- **THEN** a new row SHALL be appended with the given values

#### Scenario: Append from stdin CSV
- **GIVEN** stdin containing CSV data
- **WHEN** executing `xlex row append report.xlsx "Data" --from stdin`
- **THEN** all rows from stdin SHALL be appended

#### Scenario: Append from file
- **GIVEN** the command `xlex row append report.xlsx "Data" --from data.csv`
- **WHEN** executed
- **THEN** all rows from data.csv SHALL be appended

#### Scenario: Append with type inference
- **GIVEN** stdin containing "42,true,Hello"
- **WHEN** executing `xlex row append report.xlsx "Data" --from stdin`
- **THEN** 42 SHALL be stored as number
- **AND** true SHALL be stored as boolean
- **AND** Hello SHALL be stored as string

#### Scenario: Append preserving strings
- **GIVEN** the command with --all-strings flag
- **WHEN** executing `xlex row append report.xlsx "Data" --from stdin --all-strings`
- **THEN** all values SHALL be stored as strings

#### Scenario: Streaming append
- **GIVEN** stdin with 100,000 rows
- **WHEN** executing `xlex row append report.xlsx "Data" --from stdin`
- **THEN** memory usage SHALL remain constant
- **AND** rows SHALL be written in streaming fashion

### Requirement: Insert Rows

The system SHALL insert rows via `xlex row insert <file> <sheet> <position>`.

#### Scenario: Insert single row
- **GIVEN** the command `xlex row insert report.xlsx "Data" 5`
- **WHEN** executed
- **THEN** a new empty row SHALL be inserted at position 5
- **AND** existing rows 5+ SHALL shift down

#### Scenario: Insert multiple rows
- **GIVEN** the command `xlex row insert report.xlsx "Data" 5 --count 3`
- **WHEN** executed
- **THEN** 3 empty rows SHALL be inserted starting at position 5

#### Scenario: Insert with data
- **GIVEN** the command `xlex row insert report.xlsx "Data" 5 -- "A" "B" "C"`
- **WHEN** executed
- **THEN** a new row with values A, B, C SHALL be inserted at position 5

#### Scenario: Formula reference update
- **GIVEN** a formula "=SUM(A1:A10)" in the sheet
- **WHEN** inserting a row at position 5
- **THEN** the formula SHALL be updated to "=SUM(A1:A11)"

### Requirement: Delete Rows

The system SHALL delete rows via `xlex row delete <file> <sheet> <row>`.

#### Scenario: Delete single row
- **GIVEN** the command `xlex row delete report.xlsx "Data" 5`
- **WHEN** executed
- **THEN** row 5 SHALL be deleted
- **AND** rows 6+ SHALL shift up

#### Scenario: Delete row range
- **GIVEN** the command `xlex row delete report.xlsx "Data" 5:10`
- **WHEN** executed
- **THEN** rows 5 through 10 SHALL be deleted

#### Scenario: Delete with formula update
- **GIVEN** a formula "=SUM(A1:A10)" referencing deleted rows
- **WHEN** deleting rows 3:5
- **THEN** the formula SHALL be updated to "=SUM(A1:A7)"

#### Scenario: Delete row containing formula reference
- **GIVEN** a formula "=A5" and deleting row 5
- **WHEN** executed
- **THEN** the formula SHALL become "=#REF!"

### Requirement: Copy Rows

The system SHALL copy rows via `xlex row copy <file> <sheet> <source> <dest>`.

#### Scenario: Copy single row
- **GIVEN** the command `xlex row copy report.xlsx "Data" 5 10`
- **WHEN** executed
- **THEN** row 5 SHALL be copied to row 10
- **AND** row 10 SHALL be inserted (not overwritten)

#### Scenario: Copy row range
- **GIVEN** the command `xlex row copy report.xlsx "Data" 5:10 20`
- **WHEN** executed
- **THEN** rows 5-10 SHALL be copied starting at row 20

#### Scenario: Copy to another sheet
- **GIVEN** the command `xlex row copy report.xlsx "Data" 5 --to-sheet "Backup" 1`
- **WHEN** executed
- **THEN** row 5 from Data SHALL be copied to row 1 of Backup

#### Scenario: Copy with formula adjustment
- **GIVEN** row 5 contains "=A5+B5"
- **WHEN** copying to row 10
- **THEN** the formula SHALL become "=A10+B10"

### Requirement: Move Rows

The system SHALL move rows via `xlex row move <file> <sheet> <source> <dest>`.

#### Scenario: Move single row
- **GIVEN** the command `xlex row move report.xlsx "Data" 5 10`
- **WHEN** executed
- **THEN** row 5 SHALL be moved to position 10
- **AND** original row 5 SHALL be removed

#### Scenario: Move row range
- **GIVEN** the command `xlex row move report.xlsx "Data" 5:10 20`
- **WHEN** executed
- **THEN** rows 5-10 SHALL be moved to start at row 20

#### Scenario: Move up
- **GIVEN** the command `xlex row move report.xlsx "Data" 10 5`
- **WHEN** executed
- **THEN** row 10 SHALL be moved to position 5

### Requirement: Set Row Height

The system SHALL set row height via `xlex row height <file> <sheet> <row> <height>`.

#### Scenario: Set single row height
- **GIVEN** the command `xlex row height report.xlsx "Data" 5 30`
- **WHEN** executed
- **THEN** row 5 height SHALL be set to 30 points

#### Scenario: Set row range height
- **GIVEN** the command `xlex row height report.xlsx "Data" 5:10 25`
- **WHEN** executed
- **THEN** rows 5-10 SHALL all have height 25 points

#### Scenario: Get row height
- **GIVEN** the command `xlex row height report.xlsx "Data" 5`
- **WHEN** executed without height value
- **THEN** output SHALL be the current height of row 5

#### Scenario: Auto-fit height
- **GIVEN** the command `xlex row height report.xlsx "Data" 5 --auto`
- **WHEN** executed
- **THEN** row 5 height SHALL be set to fit content

### Requirement: Hide Rows

The system SHALL hide rows via `xlex row hide <file> <sheet> <row>`.

#### Scenario: Hide single row
- **GIVEN** the command `xlex row hide report.xlsx "Data" 5`
- **WHEN** executed
- **THEN** row 5 SHALL be hidden

#### Scenario: Hide row range
- **GIVEN** the command `xlex row hide report.xlsx "Data" 5:10`
- **WHEN** executed
- **THEN** rows 5-10 SHALL be hidden

### Requirement: Unhide Rows

The system SHALL unhide rows via `xlex row unhide <file> <sheet> <row>`.

#### Scenario: Unhide single row
- **GIVEN** the command `xlex row unhide report.xlsx "Data" 5`
- **WHEN** executed
- **THEN** row 5 SHALL be visible

#### Scenario: Unhide all rows
- **GIVEN** the command `xlex row unhide report.xlsx "Data" --all`
- **WHEN** executed
- **THEN** all hidden rows SHALL become visible

### Requirement: Find Rows

The system SHALL find rows matching criteria via `xlex row find <file> <sheet>`.

#### Scenario: Find by value
- **GIVEN** the command `xlex row find report.xlsx "Data" --value "Error"`
- **WHEN** executed
- **THEN** output SHALL list row numbers containing "Error"

#### Scenario: Find by column value
- **GIVEN** the command `xlex row find report.xlsx "Data" --column A --value "Error"`
- **WHEN** executed
- **THEN** output SHALL list row numbers where column A equals "Error"

#### Scenario: Find by regex
- **GIVEN** the command `xlex row find report.xlsx "Data" --regex "ERR-\d+"`
- **WHEN** executed
- **THEN** output SHALL list row numbers matching the pattern

#### Scenario: Find empty rows
- **GIVEN** the command `xlex row find report.xlsx "Data" --empty`
- **WHEN** executed
- **THEN** output SHALL list row numbers that are completely empty
