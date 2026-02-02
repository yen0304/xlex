# Error Handling Specification

## ADDED Requirements

### Requirement: Error Code System

The system SHALL use a structured error code system with format XLEX_EXXX.

#### Scenario: Error code format
- **GIVEN** any error occurs
- **WHEN** the error is reported
- **THEN** the error code SHALL follow format XLEX_EXXX (e.g., XLEX_E001)

#### Scenario: Error message format
- **GIVEN** any error occurs
- **WHEN** the error is reported
- **THEN** output SHALL be: `XLEX_EXXX: Human readable message`

#### Scenario: JSON error format
- **GIVEN** the flag --json-errors is set
- **WHEN** an error occurs
- **THEN** output SHALL be JSON: `{"code":"XLEX_E001","message":"...","details":{...}}`

### Requirement: File Operation Errors

The system SHALL handle file operation errors with specific codes.

#### Scenario: File not found (XLEX_E001)
- **GIVEN** a command referencing non-existent file
- **WHEN** executed
- **THEN** error code SHALL be XLEX_E001
- **AND** message SHALL include the file path

#### Scenario: File already exists (XLEX_E002)
- **GIVEN** a command creating a file that exists
- **WHEN** executed without --force
- **THEN** error code SHALL be XLEX_E002
- **AND** message SHALL suggest using --force

#### Scenario: Permission denied (XLEX_E003)
- **GIVEN** a command on a file without write permission
- **WHEN** attempting to write
- **THEN** error code SHALL be XLEX_E003
- **AND** message SHALL indicate permission issue

#### Scenario: File locked (XLEX_E004)
- **GIVEN** a file locked by another process
- **WHEN** attempting to write
- **THEN** error code SHALL be XLEX_E004
- **AND** message SHALL suggest closing other applications

#### Scenario: Disk full (XLEX_E005)
- **GIVEN** insufficient disk space
- **WHEN** attempting to write
- **THEN** error code SHALL be XLEX_E005
- **AND** message SHALL indicate space needed

### Requirement: Parse Errors

The system SHALL handle parse errors with specific codes.

#### Scenario: Invalid ZIP structure (XLEX_E011)
- **GIVEN** a file that is not a valid ZIP
- **WHEN** attempting to open
- **THEN** error code SHALL be XLEX_E011
- **AND** message SHALL indicate the file is not a valid xlsx

#### Scenario: Missing required entry (XLEX_E012)
- **GIVEN** an xlsx missing workbook.xml
- **WHEN** attempting to open
- **THEN** error code SHALL be XLEX_E012
- **AND** message SHALL specify the missing entry

#### Scenario: Malformed XML (XLEX_E013)
- **GIVEN** an xlsx with malformed XML
- **WHEN** parsing
- **THEN** error code SHALL be XLEX_E013
- **AND** message SHALL include XML location (line, column)

#### Scenario: Unsupported xlsx version (XLEX_E014)
- **GIVEN** an xlsx with unsupported features
- **WHEN** attempting to parse
- **THEN** error code SHALL be XLEX_E014
- **AND** message SHALL indicate the unsupported feature

#### Scenario: Corrupted SharedStrings (XLEX_E015)
- **GIVEN** an xlsx with corrupted SharedStrings.xml
- **WHEN** attempting to read strings
- **THEN** error code SHALL be XLEX_E015
- **AND** message SHALL indicate SharedStrings corruption

### Requirement: Reference Errors

The system SHALL handle reference errors with specific codes.

#### Scenario: Invalid cell reference (XLEX_E020)
- **GIVEN** an invalid cell reference like "1A"
- **WHEN** parsing
- **THEN** error code SHALL be XLEX_E020
- **AND** message SHALL show the invalid reference

#### Scenario: Invalid range reference (XLEX_E021)
- **GIVEN** an invalid range like "B1:A1"
- **WHEN** parsing
- **THEN** error code SHALL be XLEX_E021
- **AND** message SHALL explain the issue

#### Scenario: Reference out of bounds (XLEX_E022)
- **GIVEN** a reference beyond Excel limits (e.g., XFE1)
- **WHEN** parsing
- **THEN** error code SHALL be XLEX_E022
- **AND** message SHALL indicate the limits

### Requirement: Sheet Errors

The system SHALL handle sheet errors with specific codes.

#### Scenario: Sheet already exists (XLEX_E030)
- **GIVEN** adding a sheet with existing name
- **WHEN** executed
- **THEN** error code SHALL be XLEX_E030
- **AND** message SHALL show the duplicate name

#### Scenario: Invalid sheet name (XLEX_E031)
- **GIVEN** a sheet name with invalid characters
- **WHEN** executed
- **THEN** error code SHALL be XLEX_E031
- **AND** message SHALL list invalid characters (: \ / ? * [ ])

#### Scenario: Sheet not found (XLEX_E032)
- **GIVEN** a command referencing non-existent sheet
- **WHEN** executed
- **THEN** error code SHALL be XLEX_E032
- **AND** message SHALL list available sheets

#### Scenario: Cannot remove last sheet (XLEX_E033)
- **GIVEN** attempting to remove the only sheet
- **WHEN** executed
- **THEN** error code SHALL be XLEX_E033
- **AND** message SHALL explain workbook must have at least one sheet

#### Scenario: Invalid sheet index (XLEX_E034)
- **GIVEN** a sheet index out of range
- **WHEN** executed
- **THEN** error code SHALL be XLEX_E034
- **AND** message SHALL show valid range

#### Scenario: Cannot hide last visible (XLEX_E035)
- **GIVEN** attempting to hide the last visible sheet
- **WHEN** executed
- **THEN** error code SHALL be XLEX_E035
- **AND** message SHALL explain at least one sheet must be visible

### Requirement: Formula Errors

The system SHALL handle formula errors with specific codes.

#### Scenario: Invalid formula syntax (XLEX_E040)
- **GIVEN** a formula with syntax error
- **WHEN** setting
- **THEN** error code SHALL be XLEX_E040
- **AND** message SHALL indicate the syntax issue

#### Scenario: Circular reference (XLEX_E041)
- **GIVEN** a formula creating circular reference
- **WHEN** detected
- **THEN** error code SHALL be XLEX_E041
- **AND** message SHALL show the circular path

#### Scenario: Unknown function (XLEX_E042)
- **GIVEN** a formula with unknown function
- **WHEN** validating (if validation enabled)
- **THEN** error code SHALL be XLEX_E042
- **AND** message SHALL show the unknown function

### Requirement: Row/Column Errors

The system SHALL handle row/column errors with specific codes.

#### Scenario: Invalid row index (XLEX_E050)
- **GIVEN** a row index of 0 or negative
- **WHEN** executed
- **THEN** error code SHALL be XLEX_E050
- **AND** message SHALL indicate rows are 1-indexed

#### Scenario: Row limit exceeded (XLEX_E051)
- **GIVEN** a row index > 1048576
- **WHEN** executed
- **THEN** error code SHALL be XLEX_E051
- **AND** message SHALL show the limit

#### Scenario: Invalid column reference (XLEX_E060)
- **GIVEN** an invalid column like "123"
- **WHEN** executed
- **THEN** error code SHALL be XLEX_E060
- **AND** message SHALL show valid format

#### Scenario: Column limit exceeded (XLEX_E061)
- **GIVEN** a column beyond XFD
- **WHEN** executed
- **THEN** error code SHALL be XLEX_E061
- **AND** message SHALL show the limit

### Requirement: Import/Export Errors

The system SHALL handle import/export errors with specific codes.

#### Scenario: Invalid CSV format (XLEX_E070)
- **GIVEN** malformed CSV input
- **WHEN** importing
- **THEN** error code SHALL be XLEX_E070
- **AND** message SHALL indicate the line number

#### Scenario: Invalid JSON format (XLEX_E071)
- **GIVEN** malformed JSON input
- **WHEN** importing
- **THEN** error code SHALL be XLEX_E071
- **AND** message SHALL include JSON parse error

#### Scenario: Unsupported format (XLEX_E072)
- **GIVEN** an unsupported file format
- **WHEN** converting
- **THEN** error code SHALL be XLEX_E072
- **AND** message SHALL list supported formats

### Requirement: Style Errors

The system SHALL handle style errors with specific codes.

#### Scenario: Invalid color format (XLEX_E080)
- **GIVEN** an invalid color like "red" instead of "#FF0000"
- **WHEN** applying style
- **THEN** error code SHALL be XLEX_E080
- **AND** message SHALL show valid format

#### Scenario: Style not found (XLEX_E081)
- **GIVEN** a reference to non-existent style ID
- **WHEN** applying
- **THEN** error code SHALL be XLEX_E081
- **AND** message SHALL list available styles

#### Scenario: Invalid number format (XLEX_E082)
- **GIVEN** an invalid number format string
- **WHEN** applying
- **THEN** error code SHALL be XLEX_E082
- **AND** message SHALL show format examples

### Requirement: Error Recovery Suggestions

The system SHALL provide recovery suggestions for common errors.

#### Scenario: Suggest force flag
- **GIVEN** error XLEX_E002 (file exists)
- **WHEN** error is displayed
- **THEN** message SHALL suggest: "Use --force to overwrite"

#### Scenario: Suggest sheet names
- **GIVEN** error XLEX_E032 (sheet not found)
- **WHEN** error is displayed
- **THEN** message SHALL list available sheet names

#### Scenario: Suggest valid range
- **GIVEN** error XLEX_E022 (reference out of bounds)
- **WHEN** error is displayed
- **THEN** message SHALL show: "Valid range: A1:XFD1048576"

#### Scenario: Suggest closing applications
- **GIVEN** error XLEX_E004 (file locked)
- **WHEN** error is displayed
- **THEN** message SHALL suggest: "Close Excel or other applications using this file"

### Requirement: Error Logging

The system SHALL support error logging.

#### Scenario: Log to file
- **GIVEN** XLEX_LOG_FILE environment variable set
- **WHEN** an error occurs
- **THEN** error details SHALL be appended to the log file

#### Scenario: Debug logging
- **GIVEN** the flag --debug
- **WHEN** an error occurs
- **THEN** stack trace and context SHALL be logged

#### Scenario: Log format
- **GIVEN** logging is enabled
- **WHEN** an error is logged
- **THEN** log entry SHALL include: timestamp, error code, message, context

### Requirement: Partial Success Handling

The system SHALL handle partial success in batch operations.

#### Scenario: Batch with errors
- **GIVEN** a batch operation with some failures
- **WHEN** --continue-on-error is set
- **THEN** successful operations SHALL complete
- **AND** failures SHALL be collected and reported at end

#### Scenario: Error summary
- **GIVEN** multiple errors in batch
- **WHEN** batch completes
- **THEN** summary SHALL show: total operations, successes, failures

#### Scenario: Partial result output
- **GIVEN** a batch with partial success
- **WHEN** completed
- **THEN** exit code SHALL be non-zero
- **AND** output SHALL indicate partial completion
