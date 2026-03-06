## ADDED Requirements

### Requirement: open_workbook tool
The system SHALL provide an `open_workbook` tool that opens an xlsx file and creates a session. The tool SHALL accept a `path` parameter (required) and return a session ID, sheet count, sheet names, and file size.

#### Scenario: Open a valid xlsx file
- **WHEN** an agent calls `open_workbook` with `path` set to a valid xlsx file
- **THEN** the system returns a session ID, list of sheet names, sheet count, and file size in bytes

#### Scenario: Open an invalid file
- **WHEN** an agent calls `open_workbook` with `path` set to a non-xlsx file
- **THEN** the system returns an MCP error with `XLEX_E004` (invalid extension)

### Requirement: close_workbook tool
The system SHALL provide a `close_workbook` tool that closes an open session. The tool SHALL accept `session_id` (required) and `save` (optional boolean, default false) parameters.

#### Scenario: Close without saving
- **WHEN** an agent calls `close_workbook` with `save` set to false
- **THEN** the system removes the session and discards unsaved changes

#### Scenario: Close with save
- **WHEN** an agent calls `close_workbook` with `save` set to true
- **THEN** the system saves the workbook to its original path, then removes the session

### Requirement: workbook_info tool
The system SHALL provide a `workbook_info` tool that returns metadata for an open workbook. The tool SHALL accept `session_id` (required) and return document properties (title, creator, created date, modified date) and workbook statistics (sheet count, total cells, formula count, style count, string count, file size).

#### Scenario: Get info for an open workbook
- **WHEN** an agent calls `workbook_info` with a valid session ID
- **THEN** the system returns document properties and statistics as structured text

### Requirement: create_workbook tool
The system SHALL provide a `create_workbook` tool that creates a new empty workbook file. The tool SHALL accept `path` (required) and `sheets` (optional list of sheet names, defaults to `["Sheet1"]`) parameters. It SHALL create the file, open it as a session, and return the session ID.

#### Scenario: Create a new workbook with default sheet
- **WHEN** an agent calls `create_workbook` with only `path`
- **THEN** the system creates an xlsx file at the path with one sheet named "Sheet1" and returns a session ID

#### Scenario: Create a workbook with custom sheets
- **WHEN** an agent calls `create_workbook` with `sheets` set to `["Data", "Summary"]`
- **THEN** the system creates an xlsx file with two sheets named "Data" and "Summary"

### Requirement: save_workbook tool
The system SHALL provide a `save_workbook` tool that saves an open workbook. The tool SHALL accept `session_id` (required) and `path` (optional, for save-as) parameters.

#### Scenario: Save to original path
- **WHEN** an agent calls `save_workbook` with only `session_id`
- **THEN** the system saves the workbook to the path it was opened from

#### Scenario: Save to a new path
- **WHEN** an agent calls `save_workbook` with `path` set to a new file path
- **THEN** the system saves the workbook to the new path

### Requirement: list_sheets tool
The system SHALL provide a `list_sheets` tool that returns all sheets in a workbook. The tool SHALL accept `session_id` (required) and return each sheet's name, index, visibility, and cell count.

#### Scenario: List sheets of a workbook
- **WHEN** an agent calls `list_sheets` with a valid session ID
- **THEN** the system returns a list of sheets with name, index, visibility status, and cell count for each

### Requirement: add_sheet tool
The system SHALL provide an `add_sheet` tool that adds a new sheet to a workbook. The tool SHALL accept `session_id` (required) and `name` (required) parameters.

#### Scenario: Add a sheet with a unique name
- **WHEN** an agent calls `add_sheet` with a name that does not already exist
- **THEN** the system adds the sheet and returns the new sheet's index

#### Scenario: Add a sheet with a duplicate name
- **WHEN** an agent calls `add_sheet` with a name that already exists
- **THEN** the system returns an MCP error with `XLEX_E031`

### Requirement: remove_sheet tool
The system SHALL provide a `remove_sheet` tool that removes a sheet by name. The tool SHALL accept `session_id` (required) and `name` (required) parameters.

#### Scenario: Remove an existing sheet
- **WHEN** an agent calls `remove_sheet` with a valid sheet name and the workbook has more than one sheet
- **THEN** the system removes the sheet

#### Scenario: Remove the last remaining sheet
- **WHEN** an agent calls `remove_sheet` and the workbook has only one sheet
- **THEN** the system returns an MCP error with `XLEX_E034`

### Requirement: rename_sheet tool
The system SHALL provide a `rename_sheet` tool that renames a sheet. The tool SHALL accept `session_id` (required), `old_name` (required), and `new_name` (required) parameters.

#### Scenario: Rename a sheet
- **WHEN** an agent calls `rename_sheet` with valid old and new names
- **THEN** the system renames the sheet and confirms the change

### Requirement: read_cells tool
The system SHALL provide a `read_cells` tool that reads cell values. The tool SHALL accept `session_id` (required), `sheet` (required), and one of: `cell` (single A1 reference), `range` (A1:B2 notation), or `cells` (list of A1 references). The response SHALL include each cell's reference, value, type, and formula (if present).

#### Scenario: Read a single cell
- **WHEN** an agent calls `read_cells` with `cell` set to `"A1"`
- **THEN** the system returns the cell's reference, value, and type

#### Scenario: Read a range of cells
- **WHEN** an agent calls `read_cells` with `range` set to `"A1:C3"`
- **THEN** the system returns all cells in the range with their references, values, and types

#### Scenario: Read a list of specific cells
- **WHEN** an agent calls `read_cells` with `cells` set to `["A1", "B5", "D10"]`
- **THEN** the system returns each requested cell's reference, value, and type

#### Scenario: Read an empty cell
- **WHEN** an agent calls `read_cells` for a cell that has no value
- **THEN** the system returns the cell reference with a null/empty value and type "empty"

### Requirement: write_cells tool
The system SHALL provide a `write_cells` tool that writes values to cells. The tool SHALL accept `session_id` (required), `sheet` (required), and `cells` (required, list of `{cell, value}` objects where `cell` is A1 notation and `value` is the value to write).

#### Scenario: Write a single cell
- **WHEN** an agent calls `write_cells` with one cell entry `{cell: "A1", value: "Hello"}`
- **THEN** the system sets cell A1 to the string "Hello"

#### Scenario: Write multiple cells in batch
- **WHEN** an agent calls `write_cells` with multiple cell entries
- **THEN** the system sets all specified cells to their respective values

#### Scenario: Write a numeric value
- **WHEN** an agent calls `write_cells` with `{cell: "A1", value: 42}`
- **THEN** the system sets cell A1 to the number 42

### Requirement: clear_cells tool
The system SHALL provide a `clear_cells` tool that clears cell contents. The tool SHALL accept `session_id` (required), `sheet` (required), and `range` (required, A1:B2 notation).

#### Scenario: Clear a range of cells
- **WHEN** an agent calls `clear_cells` with range `"A1:C3"`
- **THEN** the system clears all cell values in the specified range

### Requirement: read_rows tool
The system SHALL provide a `read_rows` tool that reads rows from a sheet. The tool SHALL accept `session_id` (required), `sheet` (required), `start_row` (optional, 1-indexed, default 1), and `limit` (optional, default all rows).

#### Scenario: Read all rows
- **WHEN** an agent calls `read_rows` with only `session_id` and `sheet`
- **THEN** the system returns all rows with their cell values

#### Scenario: Read rows with limit
- **WHEN** an agent calls `read_rows` with `start_row` set to 2 and `limit` set to 10
- **THEN** the system returns rows 2 through 11

### Requirement: insert_rows tool
The system SHALL provide an `insert_rows` tool that inserts empty rows. The tool SHALL accept `session_id` (required), `sheet` (required), `row` (required, 1-indexed position), and `count` (optional, default 1).

#### Scenario: Insert rows at a position
- **WHEN** an agent calls `insert_rows` with `row` set to 5 and `count` set to 3
- **THEN** the system inserts 3 empty rows at row 5, shifting existing rows down

### Requirement: delete_rows tool
The system SHALL provide a `delete_rows` tool that deletes rows. The tool SHALL accept `session_id` (required), `sheet` (required), `row` (required, 1-indexed position), and `count` (optional, default 1).

#### Scenario: Delete rows at a position
- **WHEN** an agent calls `delete_rows` with `row` set to 5 and `count` set to 2
- **THEN** the system deletes rows 5 and 6, shifting remaining rows up

### Requirement: export_sheet tool
The system SHALL provide an `export_sheet` tool that exports sheet data. The tool SHALL accept `session_id` (required), `sheet` (required), and `format` (required, one of `csv`, `json`, `markdown`). The tool SHALL return the exported content as text.

#### Scenario: Export as CSV
- **WHEN** an agent calls `export_sheet` with `format` set to `"csv"`
- **THEN** the system returns the sheet data as a CSV string

#### Scenario: Export as JSON
- **WHEN** an agent calls `export_sheet` with `format` set to `"json"`
- **THEN** the system returns the sheet data as a JSON array of row objects

#### Scenario: Export as Markdown
- **WHEN** an agent calls `export_sheet` with `format` set to `"markdown"`
- **THEN** the system returns the sheet data as a markdown table
