# Import/Export Specification

## ADDED Requirements

### Requirement: Export to CSV

The system SHALL export sheet data to CSV via `xlex to csv <file> <sheet>`.

#### Scenario: Export entire sheet
- **GIVEN** the command `xlex to csv report.xlsx "Data"`
- **WHEN** executed
- **THEN** output SHALL be CSV format to stdout

#### Scenario: Export to file
- **GIVEN** the command `xlex to csv report.xlsx "Data" --output data.csv`
- **WHEN** executed
- **THEN** data.csv SHALL be created with CSV content

#### Scenario: Export range
- **GIVEN** the command `xlex to csv report.xlsx "Data" --range A1:D100`
- **WHEN** executed
- **THEN** only the specified range SHALL be exported

#### Scenario: Custom delimiter
- **GIVEN** the command `xlex to csv report.xlsx "Data" --delimiter ";"`
- **WHEN** executed
- **THEN** output SHALL use semicolon as delimiter

#### Scenario: Custom quote character
- **GIVEN** the command `xlex to csv report.xlsx "Data" --quote "'"`
- **WHEN** executed
- **THEN** output SHALL use single quote for quoting

#### Scenario: No header
- **GIVEN** the command `xlex to csv report.xlsx "Data" --no-header`
- **WHEN** executed
- **THEN** first row SHALL not be treated specially

#### Scenario: Export formulas
- **GIVEN** the command `xlex to csv report.xlsx "Data" --formulas`
- **WHEN** executed
- **THEN** formulas SHALL be exported instead of values

#### Scenario: Streaming export
- **GIVEN** a sheet with 1,000,000 rows
- **WHEN** executing `xlex to csv report.xlsx "Data"`
- **THEN** memory usage SHALL remain constant
- **AND** output SHALL stream to stdout

### Requirement: Export to JSON

The system SHALL export sheet data to JSON via `xlex to json <file> <sheet>`.

#### Scenario: Export as array of arrays
- **GIVEN** the command `xlex to json report.xlsx "Data"`
- **WHEN** executed
- **THEN** output SHALL be 2D JSON array

#### Scenario: Export as records
- **GIVEN** the command `xlex to json report.xlsx "Data" --records`
- **WHEN** row 1 contains headers
- **THEN** output SHALL be JSON array of objects

#### Scenario: Export to file
- **GIVEN** the command `xlex to json report.xlsx "Data" --output data.json`
- **WHEN** executed
- **THEN** data.json SHALL be created

#### Scenario: Pretty print
- **GIVEN** the command `xlex to json report.xlsx "Data" --pretty`
- **WHEN** executed
- **THEN** output SHALL be formatted with indentation

#### Scenario: Export range
- **GIVEN** the command `xlex to json report.xlsx "Data" --range A1:D100`
- **WHEN** executed
- **THEN** only the specified range SHALL be exported

#### Scenario: Include metadata
- **GIVEN** the command `xlex to json report.xlsx "Data" --with-metadata`
- **WHEN** executed
- **THEN** output SHALL include sheet name, dimensions, and export timestamp

#### Scenario: Custom null handling
- **GIVEN** the command `xlex to json report.xlsx "Data" --null-value ""`
- **WHEN** executed
- **THEN** empty cells SHALL be represented as empty strings

### Requirement: Export to NDJSON

The system SHALL export sheet data to NDJSON via `xlex to ndjson <file> <sheet>`.

#### Scenario: Export streaming
- **GIVEN** the command `xlex to ndjson report.xlsx "Data"`
- **WHEN** executed
- **THEN** output SHALL be one JSON object per line

#### Scenario: With headers
- **GIVEN** the command `xlex to ndjson report.xlsx "Data" --headers`
- **WHEN** row 1 contains headers
- **THEN** each line SHALL be an object with header keys

#### Scenario: Without headers
- **GIVEN** the command `xlex to ndjson report.xlsx "Data"`
- **WHEN** executed without --headers
- **THEN** each line SHALL be a JSON array

#### Scenario: Large file streaming
- **GIVEN** a sheet with 1,000,000 rows
- **WHEN** executing `xlex to ndjson report.xlsx "Data"`
- **THEN** memory usage SHALL remain constant
- **AND** output SHALL stream line by line

### Requirement: Import from CSV

The system SHALL import CSV data via `xlex from csv <file> <sheet>`.

#### Scenario: Import from stdin
- **GIVEN** the command `cat data.csv | xlex from csv report.xlsx "Data"`
- **WHEN** executed
- **THEN** CSV data SHALL be imported to the sheet

#### Scenario: Import from file
- **GIVEN** the command `xlex from csv report.xlsx "Data" --input data.csv`
- **WHEN** executed
- **THEN** data.csv content SHALL be imported

#### Scenario: Import to new sheet
- **GIVEN** the command `xlex from csv report.xlsx "NewSheet" --input data.csv`
- **WHEN** NewSheet doesn't exist
- **THEN** NewSheet SHALL be created with the data

#### Scenario: Import with append
- **GIVEN** the command `xlex from csv report.xlsx "Data" --input data.csv --append`
- **WHEN** executed
- **THEN** data SHALL be appended after existing rows

#### Scenario: Import with replace
- **GIVEN** the command `xlex from csv report.xlsx "Data" --input data.csv --replace`
- **WHEN** executed
- **THEN** existing data SHALL be cleared before import

#### Scenario: Custom delimiter
- **GIVEN** the command `xlex from csv report.xlsx "Data" --input data.csv --delimiter ";"`
- **WHEN** executed
- **THEN** semicolon SHALL be used as delimiter

#### Scenario: Type inference
- **GIVEN** CSV with values "42", "true", "Hello"
- **WHEN** importing without flags
- **THEN** 42 SHALL be stored as number
- **AND** true SHALL be stored as boolean
- **AND** Hello SHALL be stored as string

#### Scenario: All as strings
- **GIVEN** the command `xlex from csv report.xlsx "Data" --input data.csv --all-strings`
- **WHEN** executed
- **THEN** all values SHALL be stored as strings

#### Scenario: First row as header
- **GIVEN** the command `xlex from csv report.xlsx "Data" --input data.csv --header`
- **WHEN** executed
- **THEN** first row SHALL be treated as header
- **AND** header style SHALL be applied

#### Scenario: Streaming import
- **GIVEN** a CSV with 1,000,000 rows
- **WHEN** importing
- **THEN** memory usage SHALL remain constant
- **AND** rows SHALL be written in streaming fashion

### Requirement: Import from JSON

The system SHALL import JSON data via `xlex from json <file> <sheet>`.

#### Scenario: Import array of arrays
- **GIVEN** JSON input `[[1,2,3],[4,5,6]]`
- **WHEN** executing `xlex from json report.xlsx "Data"`
- **THEN** data SHALL be imported as 2 rows, 3 columns

#### Scenario: Import array of objects
- **GIVEN** JSON input `[{"name":"John","age":30}]`
- **WHEN** executing `xlex from json report.xlsx "Data"`
- **THEN** headers SHALL be created from object keys
- **AND** values SHALL be in subsequent rows

#### Scenario: Import from file
- **GIVEN** the command `xlex from json report.xlsx "Data" --input data.json`
- **WHEN** executed
- **THEN** data.json content SHALL be imported

#### Scenario: Import with path
- **GIVEN** JSON input `{"data":{"rows":[[1,2,3]]}}`
- **WHEN** executing `xlex from json report.xlsx "Data" --path "data.rows"`
- **THEN** only the nested array SHALL be imported

#### Scenario: Flatten nested objects
- **GIVEN** JSON with nested objects
- **WHEN** executing `xlex from json report.xlsx "Data" --flatten`
- **THEN** nested keys SHALL be flattened (e.g., "address.city")

### Requirement: Import from NDJSON

The system SHALL import NDJSON data via `xlex from ndjson <file> <sheet>`.

#### Scenario: Import streaming
- **GIVEN** NDJSON input from stdin
- **WHEN** executing `xlex from ndjson report.xlsx "Data"`
- **THEN** each line SHALL become a row

#### Scenario: Import objects
- **GIVEN** NDJSON with objects `{"a":1}\n{"a":2}`
- **WHEN** executing `xlex from ndjson report.xlsx "Data"`
- **THEN** headers SHALL be created from first object's keys

#### Scenario: Import arrays
- **GIVEN** NDJSON with arrays `[1,2]\n[3,4]`
- **WHEN** executing `xlex from ndjson report.xlsx "Data"`
- **THEN** each array SHALL become a row

#### Scenario: Large file streaming
- **GIVEN** NDJSON with 1,000,000 lines
- **WHEN** importing
- **THEN** memory usage SHALL remain constant

### Requirement: Export Multiple Sheets

The system SHALL export multiple sheets via `xlex to <format> <file> --all`.

#### Scenario: Export all sheets to CSV
- **GIVEN** the command `xlex to csv report.xlsx --all --output-dir ./export/`
- **WHEN** executed
- **THEN** each sheet SHALL be exported to a separate CSV file

#### Scenario: Export all sheets to JSON
- **GIVEN** the command `xlex to json report.xlsx --all`
- **WHEN** executed
- **THEN** output SHALL be JSON object with sheet names as keys

### Requirement: Export Workbook Metadata

The system SHALL export workbook metadata via `xlex to meta <file>`.

#### Scenario: Export metadata
- **GIVEN** the command `xlex to meta report.xlsx`
- **WHEN** executed
- **THEN** output SHALL include:
  - Document properties
  - Sheet list with dimensions
  - Named ranges
  - Style count

#### Scenario: JSON metadata
- **GIVEN** the command `xlex to meta report.xlsx --format json`
- **WHEN** executed
- **THEN** output SHALL be valid JSON

### Requirement: Database Pipeline Integration

The system SHALL support database pipeline integration.

#### Scenario: PostgreSQL import
- **GIVEN** the command `psql -c "COPY table TO STDOUT CSV" | xlex from csv report.xlsx "Data"`
- **WHEN** executed
- **THEN** database output SHALL be imported

#### Scenario: SQLite export
- **GIVEN** the command `xlex to csv report.xlsx "Data" | sqlite3 db.sqlite ".import /dev/stdin table"`
- **WHEN** executed
- **THEN** sheet data SHALL be imported to SQLite

#### Scenario: MySQL import
- **GIVEN** the command `mysql -e "SELECT * FROM table" --batch | xlex from csv report.xlsx "Data" --delimiter "\t"`
- **WHEN** executed
- **THEN** MySQL output SHALL be imported

### Requirement: Format Conversion

The system SHALL convert between formats via `xlex convert`.

#### Scenario: CSV to XLSX
- **GIVEN** the command `xlex convert data.csv data.xlsx`
- **WHEN** executed
- **THEN** data.xlsx SHALL be created from CSV

#### Scenario: JSON to XLSX
- **GIVEN** the command `xlex convert data.json data.xlsx`
- **WHEN** executed
- **THEN** data.xlsx SHALL be created from JSON

#### Scenario: XLSX to CSV
- **GIVEN** the command `xlex convert report.xlsx report.csv`
- **WHEN** executed
- **THEN** first sheet SHALL be exported to CSV

#### Scenario: Multiple sheets to CSV
- **GIVEN** the command `xlex convert report.xlsx ./output/ --all`
- **WHEN** executed
- **THEN** each sheet SHALL become a separate CSV file
