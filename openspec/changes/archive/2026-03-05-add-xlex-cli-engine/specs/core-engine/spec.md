# Core Engine Specification

## ADDED Requirements

### Requirement: Streaming ZIP Archive Access

The system SHALL provide streaming access to ZIP archive entries without loading the entire archive into memory.

#### Scenario: Open large xlsx file
- **GIVEN** an xlsx file of 200MB
- **WHEN** the system opens the file
- **THEN** memory usage SHALL remain under 50MB
- **AND** the file SHALL be accessible within 100ms

#### Scenario: Access specific ZIP entry
- **GIVEN** an open xlsx archive
- **WHEN** requesting a specific entry (e.g., xl/worksheets/sheet1.xml)
- **THEN** only that entry SHALL be read from disk
- **AND** other entries SHALL remain unread

### Requirement: SAX-based XML Parsing

The system SHALL use event-based (SAX) XML parsing for all xlsx XML content.

#### Scenario: Parse sheet with million rows
- **GIVEN** a sheet XML with 1,000,000 rows
- **WHEN** streaming through the sheet
- **THEN** memory usage SHALL remain constant regardless of row count
- **AND** each row SHALL be yielded as an iterator item

#### Scenario: Parse malformed XML
- **GIVEN** an xlsx with malformed XML content
- **WHEN** attempting to parse
- **THEN** the system SHALL return error code XLEX_E010 (ParseError)
- **AND** the error message SHALL indicate the XML location

### Requirement: Lazy SharedStrings Loading

The system SHALL load SharedStrings entries on-demand with LRU caching.

#### Scenario: Access string by index
- **GIVEN** an xlsx with 100,000 unique strings
- **WHEN** accessing string at index 50,000
- **THEN** only necessary portions of SharedStrings.xml SHALL be parsed
- **AND** the string SHALL be cached for subsequent access

#### Scenario: LRU cache eviction
- **GIVEN** an LRU cache with 10,000 entry limit
- **WHEN** accessing the 10,001st unique string
- **THEN** the least recently used entry SHALL be evicted
- **AND** cache size SHALL not exceed the configured limit

#### Scenario: Configure cache size
- **GIVEN** the --string-cache-size flag set to 5000
- **WHEN** the system initializes
- **THEN** the LRU cache limit SHALL be 5000 entries

### Requirement: Style Registry Management

The system SHALL maintain a registry of cell styles parsed from styles.xml.

#### Scenario: Load style definitions
- **GIVEN** an xlsx with custom styles
- **WHEN** opening the workbook
- **THEN** style definitions SHALL be indexed by style ID
- **AND** styles SHALL be accessible for cell formatting queries

#### Scenario: Style lookup by ID
- **GIVEN** a cell with style_id=5
- **WHEN** querying the cell's style
- **THEN** the system SHALL return the style definition for ID 5
- **AND** the style SHALL include font, fill, border, and number format

### Requirement: Copy-on-Write ZIP Modification

The system SHALL use copy-on-write strategy when modifying xlsx files.

#### Scenario: Modify single cell
- **GIVEN** an xlsx with 100 sheets
- **WHEN** modifying one cell in sheet 1
- **THEN** only sheet1.xml SHALL be rewritten
- **AND** all other ZIP entries SHALL be copied unchanged

#### Scenario: Atomic write completion
- **GIVEN** a modification operation in progress
- **WHEN** the operation completes successfully
- **THEN** the output file SHALL be atomically renamed from temp
- **AND** the original file SHALL remain unchanged until rename

#### Scenario: Write failure recovery
- **GIVEN** a modification operation in progress
- **WHEN** the operation fails (e.g., disk full)
- **THEN** the original file SHALL remain unchanged
- **AND** the temp file SHALL be cleaned up

### Requirement: Cell Reference Parsing

The system SHALL parse A1-style cell references.

#### Scenario: Parse simple reference
- **GIVEN** the reference "A1"
- **WHEN** parsing the reference
- **THEN** column SHALL be 1 (A)
- **AND** row SHALL be 1

#### Scenario: Parse multi-letter column
- **GIVEN** the reference "AA100"
- **WHEN** parsing the reference
- **THEN** column SHALL be 27 (AA)
- **AND** row SHALL be 100

#### Scenario: Parse maximum reference
- **GIVEN** the reference "XFD1048576"
- **WHEN** parsing the reference
- **THEN** column SHALL be 16384 (XFD)
- **AND** row SHALL be 1048576

#### Scenario: Invalid reference format
- **GIVEN** the reference "1A" or "A0"
- **WHEN** parsing the reference
- **THEN** the system SHALL return error code XLEX_E020 (InvalidReference)

### Requirement: Range Reference Parsing

The system SHALL parse A1-style range references.

#### Scenario: Parse simple range
- **GIVEN** the range "A1:B10"
- **WHEN** parsing the range
- **THEN** start SHALL be A1
- **AND** end SHALL be B10

#### Scenario: Parse full column range
- **GIVEN** the range "A:A"
- **WHEN** parsing the range
- **THEN** the range SHALL represent all cells in column A

#### Scenario: Parse full row range
- **GIVEN** the range "1:1"
- **WHEN** parsing the range
- **THEN** the range SHALL represent all cells in row 1

#### Scenario: Invalid range format
- **GIVEN** the range "B1:A1" (end before start)
- **WHEN** parsing the range
- **THEN** the system SHALL return error code XLEX_E021 (InvalidRange)

### Requirement: Workbook Relationship Tracking

The system SHALL track relationships between workbook components.

#### Scenario: Resolve sheet file path
- **GIVEN** a sheet named "Data"
- **WHEN** querying the sheet's XML path
- **THEN** the system SHALL resolve via workbook.xml.rels
- **AND** return the correct path (e.g., xl/worksheets/sheet3.xml)

#### Scenario: Track new sheet relationships
- **GIVEN** a new sheet being added
- **WHEN** the sheet is created
- **THEN** a new relationship entry SHALL be added to workbook.xml.rels
- **AND** the relationship ID SHALL be unique
