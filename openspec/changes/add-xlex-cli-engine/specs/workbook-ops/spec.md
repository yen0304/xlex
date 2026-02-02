# Workbook Operations Specification

## ADDED Requirements

### Requirement: Display Workbook Information

The system SHALL display comprehensive workbook metadata via `xlex info <file>`.

#### Scenario: Show basic info
- **GIVEN** the command `xlex info report.xlsx`
- **WHEN** executed
- **THEN** output SHALL include:
  - File name and size
  - Number of sheets
  - Sheet names
  - Created/modified dates
  - Author (if present)

#### Scenario: JSON output format
- **GIVEN** the command `xlex info report.xlsx --format json`
- **WHEN** executed
- **THEN** output SHALL be valid JSON with all metadata fields

#### Scenario: File not found
- **GIVEN** the command `xlex info nonexistent.xlsx`
- **WHEN** executed
- **THEN** exit code SHALL be non-zero
- **AND** error code SHALL be XLEX_E001 (FileNotFound)

### Requirement: Validate Workbook Structure

The system SHALL validate xlsx file structure via `xlex validate <file>`.

#### Scenario: Valid workbook
- **GIVEN** the command `xlex validate valid.xlsx`
- **WHEN** executed on a well-formed xlsx
- **THEN** exit code SHALL be 0
- **AND** output SHALL indicate "Valid"

#### Scenario: Corrupted ZIP structure
- **GIVEN** the command `xlex validate corrupted.xlsx`
- **WHEN** executed on a file with invalid ZIP structure
- **THEN** exit code SHALL be non-zero
- **AND** error code SHALL be XLEX_E011 (InvalidZipStructure)

#### Scenario: Missing required entries
- **GIVEN** an xlsx missing workbook.xml
- **WHEN** validated
- **THEN** error code SHALL be XLEX_E012 (MissingRequiredEntry)
- **AND** message SHALL specify the missing entry

#### Scenario: Verbose validation
- **GIVEN** the command `xlex validate file.xlsx --verbose`
- **WHEN** executed
- **THEN** output SHALL list all validated components
- **AND** include warnings for non-critical issues

### Requirement: Clone Workbook

The system SHALL create exact copies of workbooks via `xlex clone <src> <dest>`.

#### Scenario: Simple clone
- **GIVEN** the command `xlex clone source.xlsx copy.xlsx`
- **WHEN** executed
- **THEN** copy.xlsx SHALL be byte-identical to source.xlsx
- **AND** source.xlsx SHALL remain unchanged

#### Scenario: Clone to existing file
- **GIVEN** the command `xlex clone source.xlsx existing.xlsx`
- **WHEN** existing.xlsx already exists
- **THEN** exit code SHALL be non-zero
- **AND** error code SHALL be XLEX_E002 (FileExists)

#### Scenario: Force overwrite
- **GIVEN** the command `xlex clone source.xlsx existing.xlsx --force`
- **WHEN** existing.xlsx already exists
- **THEN** existing.xlsx SHALL be overwritten
- **AND** exit code SHALL be 0

### Requirement: Create New Workbook

The system SHALL create new empty workbooks via `xlex create <file>`.

#### Scenario: Create empty workbook
- **GIVEN** the command `xlex create new.xlsx`
- **WHEN** executed
- **THEN** new.xlsx SHALL be created
- **AND** SHALL contain one sheet named "Sheet1"
- **AND** SHALL be a valid xlsx file

#### Scenario: Create with custom sheet name
- **GIVEN** the command `xlex create new.xlsx --sheet "Data"`
- **WHEN** executed
- **THEN** new.xlsx SHALL contain one sheet named "Data"

#### Scenario: Create with multiple sheets
- **GIVEN** the command `xlex create new.xlsx --sheets "Data,Summary,Config"`
- **WHEN** executed
- **THEN** new.xlsx SHALL contain three sheets in order: Data, Summary, Config

#### Scenario: File already exists
- **GIVEN** the command `xlex create existing.xlsx`
- **WHEN** existing.xlsx already exists
- **THEN** exit code SHALL be non-zero
- **AND** error code SHALL be XLEX_E002 (FileExists)

### Requirement: Get Workbook Properties

The system SHALL retrieve document properties via `xlex props get <file> [property]`.

#### Scenario: Get all properties
- **GIVEN** the command `xlex props get report.xlsx`
- **WHEN** executed
- **THEN** output SHALL list all document properties:
  - title, subject, creator, keywords, description
  - lastModifiedBy, created, modified

#### Scenario: Get specific property
- **GIVEN** the command `xlex props get report.xlsx title`
- **WHEN** executed
- **THEN** output SHALL be only the title value

#### Scenario: Property not set
- **GIVEN** the command `xlex props get report.xlsx keywords`
- **WHEN** keywords property is not set
- **THEN** output SHALL be empty
- **AND** exit code SHALL be 0

### Requirement: Set Workbook Properties

The system SHALL modify document properties via `xlex props set <file> <property> <value>`.

#### Scenario: Set title
- **GIVEN** the command `xlex props set report.xlsx title "Q4 Report"`
- **WHEN** executed
- **THEN** the title property SHALL be set to "Q4 Report"
- **AND** the file SHALL be modified in place

#### Scenario: Set multiple properties
- **GIVEN** the command `xlex props set report.xlsx --title "Report" --author "John"`
- **WHEN** executed
- **THEN** both properties SHALL be updated

#### Scenario: Clear property
- **GIVEN** the command `xlex props set report.xlsx keywords ""`
- **WHEN** executed
- **THEN** the keywords property SHALL be cleared

#### Scenario: Output to different file
- **GIVEN** the command `xlex props set report.xlsx title "New" --output modified.xlsx`
- **WHEN** executed
- **THEN** modified.xlsx SHALL have the new title
- **AND** report.xlsx SHALL remain unchanged

### Requirement: Workbook Statistics

The system SHALL provide detailed statistics via `xlex stats <file>`.

#### Scenario: Show statistics
- **GIVEN** the command `xlex stats report.xlsx`
- **WHEN** executed
- **THEN** output SHALL include:
  - Total cell count
  - Non-empty cell count
  - Formula count
  - Unique string count
  - Style count
  - Per-sheet breakdown

#### Scenario: JSON statistics
- **GIVEN** the command `xlex stats report.xlsx --format json`
- **WHEN** executed
- **THEN** output SHALL be valid JSON with all statistics
