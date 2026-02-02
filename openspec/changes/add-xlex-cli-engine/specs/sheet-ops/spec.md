# Sheet Operations Specification

## ADDED Requirements

### Requirement: List Sheets

The system SHALL list all sheets in a workbook via `xlex sheet list <file>`.

#### Scenario: List sheets simple
- **GIVEN** the command `xlex sheet list report.xlsx`
- **WHEN** executed on a workbook with sheets "Data", "Summary", "Config"
- **THEN** output SHALL list sheet names in order, one per line

#### Scenario: List sheets with details
- **GIVEN** the command `xlex sheet list report.xlsx --long`
- **WHEN** executed
- **THEN** output SHALL include for each sheet:
  - Index (0-based)
  - Name
  - Visibility (visible/hidden/veryHidden)
  - Dimension (e.g., A1:Z100)
  - Row count
  - Column count

#### Scenario: JSON output
- **GIVEN** the command `xlex sheet list report.xlsx --format json`
- **WHEN** executed
- **THEN** output SHALL be a JSON array of sheet objects

#### Scenario: Empty workbook
- **GIVEN** a workbook with no sheets (edge case)
- **WHEN** listing sheets
- **THEN** output SHALL be empty
- **AND** exit code SHALL be 0

### Requirement: Add Sheet

The system SHALL add new sheets via `xlex sheet add <file> <name>`.

#### Scenario: Add sheet at end
- **GIVEN** the command `xlex sheet add report.xlsx "NewSheet"`
- **WHEN** executed
- **THEN** "NewSheet" SHALL be added as the last sheet
- **AND** the sheet SHALL be empty

#### Scenario: Add sheet at position
- **GIVEN** the command `xlex sheet add report.xlsx "NewSheet" --at 0`
- **WHEN** executed
- **THEN** "NewSheet" SHALL be inserted at position 0 (first)

#### Scenario: Duplicate sheet name
- **GIVEN** the command `xlex sheet add report.xlsx "Data"`
- **WHEN** "Data" sheet already exists
- **THEN** exit code SHALL be non-zero
- **AND** error code SHALL be XLEX_E030 (SheetExists)

#### Scenario: Invalid sheet name
- **GIVEN** the command `xlex sheet add report.xlsx "Sheet:1"`
- **WHEN** executed (colon is invalid in sheet names)
- **THEN** error code SHALL be XLEX_E031 (InvalidSheetName)

#### Scenario: Sheet name too long
- **GIVEN** a sheet name longer than 31 characters
- **WHEN** attempting to add
- **THEN** error code SHALL be XLEX_E031 (InvalidSheetName)
- **AND** message SHALL indicate the 31 character limit

### Requirement: Remove Sheet

The system SHALL remove sheets via `xlex sheet remove <file> <name>`.

#### Scenario: Remove existing sheet
- **GIVEN** the command `xlex sheet remove report.xlsx "OldSheet"`
- **WHEN** executed
- **THEN** "OldSheet" SHALL be removed from the workbook
- **AND** all references to the sheet SHALL be cleaned up

#### Scenario: Remove non-existent sheet
- **GIVEN** the command `xlex sheet remove report.xlsx "NoSuchSheet"`
- **WHEN** executed
- **THEN** error code SHALL be XLEX_E032 (SheetNotFound)

#### Scenario: Remove last sheet
- **GIVEN** a workbook with only one sheet
- **WHEN** attempting to remove it
- **THEN** error code SHALL be XLEX_E033 (CannotRemoveLastSheet)

#### Scenario: Remove by index
- **GIVEN** the command `xlex sheet remove report.xlsx --index 2`
- **WHEN** executed
- **THEN** the sheet at index 2 SHALL be removed

### Requirement: Rename Sheet

The system SHALL rename sheets via `xlex sheet rename <file> <old> <new>`.

#### Scenario: Rename sheet
- **GIVEN** the command `xlex sheet rename report.xlsx "Sheet1" "Data"`
- **WHEN** executed
- **THEN** the sheet SHALL be renamed from "Sheet1" to "Data"
- **AND** all formula references SHALL be updated

#### Scenario: Rename to existing name
- **GIVEN** the command `xlex sheet rename report.xlsx "Sheet1" "Sheet2"`
- **WHEN** "Sheet2" already exists
- **THEN** error code SHALL be XLEX_E030 (SheetExists)

#### Scenario: Rename non-existent sheet
- **GIVEN** the command `xlex sheet rename report.xlsx "NoSheet" "New"`
- **WHEN** executed
- **THEN** error code SHALL be XLEX_E032 (SheetNotFound)

### Requirement: Copy Sheet

The system SHALL copy sheets via `xlex sheet copy <file> <source> <dest>`.

#### Scenario: Copy within workbook
- **GIVEN** the command `xlex sheet copy report.xlsx "Template" "Copy"`
- **WHEN** executed
- **THEN** a new sheet "Copy" SHALL be created
- **AND** SHALL contain all data and formatting from "Template"

#### Scenario: Copy to another workbook
- **GIVEN** the command `xlex sheet copy source.xlsx "Data" --to target.xlsx`
- **WHEN** executed
- **THEN** "Data" sheet SHALL be copied to target.xlsx
- **AND** source.xlsx SHALL remain unchanged

#### Scenario: Copy with rename
- **GIVEN** the command `xlex sheet copy report.xlsx "Data" "Data_Backup"`
- **WHEN** executed
- **THEN** the copy SHALL be named "Data_Backup"

### Requirement: Move Sheet

The system SHALL reorder sheets via `xlex sheet move <file> <name> <position>`.

#### Scenario: Move to beginning
- **GIVEN** the command `xlex sheet move report.xlsx "Summary" 0`
- **WHEN** executed
- **THEN** "Summary" SHALL become the first sheet (index 0)

#### Scenario: Move to end
- **GIVEN** the command `xlex sheet move report.xlsx "Data" -1`
- **WHEN** executed (-1 means last position)
- **THEN** "Data" SHALL become the last sheet

#### Scenario: Move to specific position
- **GIVEN** the command `xlex sheet move report.xlsx "Config" 2`
- **WHEN** executed
- **THEN** "Config" SHALL be at index 2

#### Scenario: Invalid position
- **GIVEN** the command `xlex sheet move report.xlsx "Data" 100`
- **WHEN** there are only 3 sheets
- **THEN** error code SHALL be XLEX_E034 (InvalidSheetIndex)

### Requirement: Hide Sheet

The system SHALL hide sheets via `xlex sheet hide <file> <name>`.

#### Scenario: Hide sheet
- **GIVEN** the command `xlex sheet hide report.xlsx "Internal"`
- **WHEN** executed
- **THEN** "Internal" sheet visibility SHALL be set to "hidden"

#### Scenario: Very hidden
- **GIVEN** the command `xlex sheet hide report.xlsx "Secret" --very`
- **WHEN** executed
- **THEN** "Internal" sheet visibility SHALL be set to "veryHidden"
- **AND** the sheet SHALL not be visible in Excel's unhide dialog

#### Scenario: Hide last visible sheet
- **GIVEN** all other sheets are hidden
- **WHEN** attempting to hide the last visible sheet
- **THEN** error code SHALL be XLEX_E035 (CannotHideLastVisible)

### Requirement: Unhide Sheet

The system SHALL unhide sheets via `xlex sheet unhide <file> <name>`.

#### Scenario: Unhide hidden sheet
- **GIVEN** the command `xlex sheet unhide report.xlsx "Internal"`
- **WHEN** executed on a hidden sheet
- **THEN** the sheet visibility SHALL be set to "visible"

#### Scenario: Unhide very hidden sheet
- **GIVEN** the command `xlex sheet unhide report.xlsx "Secret"`
- **WHEN** executed on a veryHidden sheet
- **THEN** the sheet visibility SHALL be set to "visible"

#### Scenario: Unhide all
- **GIVEN** the command `xlex sheet unhide report.xlsx --all`
- **WHEN** executed
- **THEN** all hidden and veryHidden sheets SHALL become visible

### Requirement: Sheet Information

The system SHALL display detailed sheet info via `xlex sheet info <file> <name>`.

#### Scenario: Show sheet info
- **GIVEN** the command `xlex sheet info report.xlsx "Data"`
- **WHEN** executed
- **THEN** output SHALL include:
  - Sheet name and index
  - Dimension (used range)
  - Row count and column count
  - Visibility state
  - Tab color (if set)
  - Protection status
  - Freeze pane settings

#### Scenario: JSON output
- **GIVEN** the command `xlex sheet info report.xlsx "Data" --format json`
- **WHEN** executed
- **THEN** output SHALL be valid JSON with all sheet metadata

### Requirement: Set Active Sheet

The system SHALL set the active sheet via `xlex sheet active <file> <name>`.

#### Scenario: Set active sheet
- **GIVEN** the command `xlex sheet active report.xlsx "Summary"`
- **WHEN** executed
- **THEN** "Summary" SHALL be the active sheet when opened in Excel

#### Scenario: Get active sheet
- **GIVEN** the command `xlex sheet active report.xlsx`
- **WHEN** executed without a sheet name
- **THEN** output SHALL be the name of the currently active sheet
