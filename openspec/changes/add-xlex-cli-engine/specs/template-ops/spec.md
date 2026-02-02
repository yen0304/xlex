# Template Operations Specification

## ADDED Requirements

### Requirement: Apply Template

The system SHALL apply data to templates via `xlex template apply <template> <data>`.

#### Scenario: Apply JSON data
- **GIVEN** the command `xlex template apply template.xlsx data.json --output report.xlsx`
- **WHEN** template.xlsx contains placeholders like `{{name}}`, `{{date}}`
- **AND** data.json contains `{"name": "John", "date": "2024-01-15"}`
- **THEN** report.xlsx SHALL have placeholders replaced with values

#### Scenario: Apply CSV data
- **GIVEN** the command `xlex template apply template.xlsx data.csv --output report.xlsx`
- **WHEN** executed
- **THEN** CSV data SHALL be used to fill the template

#### Scenario: Apply from stdin
- **GIVEN** the command `cat data.json | xlex template apply template.xlsx - --output report.xlsx`
- **WHEN** executed
- **THEN** stdin JSON SHALL be used as data source

#### Scenario: Multiple rows
- **GIVEN** data.json containing an array of objects
- **WHEN** template has a row marked with `{{#each items}}`
- **THEN** the row SHALL be duplicated for each item in the array

#### Scenario: Preserve formatting
- **GIVEN** a template with styled cells containing placeholders
- **WHEN** applying data
- **THEN** cell formatting SHALL be preserved after replacement

#### Scenario: Missing placeholder data
- **GIVEN** a placeholder `{{missing}}` with no corresponding data
- **WHEN** applying template
- **THEN** behavior SHALL depend on --missing flag:
  - `--missing keep` (default): Keep placeholder text
  - `--missing empty`: Replace with empty string
  - `--missing error`: Exit with error

#### Scenario: Nested data
- **GIVEN** data.json with nested objects `{"user": {"name": "John"}}`
- **WHEN** template contains `{{user.name}}`
- **THEN** the nested value SHALL be resolved

#### Scenario: Date formatting
- **GIVEN** data with ISO date `{"date": "2024-01-15"}`
- **WHEN** template contains `{{date|date:YYYY-MM-DD}}`
- **THEN** the date SHALL be formatted accordingly

#### Scenario: Number formatting
- **GIVEN** data with number `{"amount": 1234.5}`
- **WHEN** template contains `{{amount|number:#,##0.00}}`
- **THEN** the number SHALL be formatted as "1,234.50"

### Requirement: Template Syntax

The system SHALL support Handlebars-like template syntax.

#### Scenario: Simple placeholder
- **GIVEN** template cell with `{{fieldName}}`
- **WHEN** data contains `{"fieldName": "value"}`
- **THEN** `{{fieldName}}` SHALL be replaced with "value"

#### Scenario: Each loop
- **GIVEN** template with row containing `{{#each items}}...{{/each}}`
- **WHEN** data contains `{"items": [{...}, {...}]}`
- **THEN** the row SHALL be duplicated for each item

#### Scenario: Conditional
- **GIVEN** template with `{{#if condition}}value{{/if}}`
- **WHEN** condition is truthy
- **THEN** "value" SHALL be shown

#### Scenario: Conditional else
- **GIVEN** template with `{{#if condition}}yes{{else}}no{{/if}}`
- **WHEN** condition is falsy
- **THEN** "no" SHALL be shown

#### Scenario: Index in loop
- **GIVEN** template with `{{#each items}}{{@index}}{{/each}}`
- **WHEN** iterating
- **THEN** `{{@index}}` SHALL be replaced with 0-based index

#### Scenario: First/Last in loop
- **GIVEN** template with `{{#each items}}{{#if @first}}First{{/if}}{{/each}}`
- **WHEN** iterating
- **THEN** "First" SHALL only appear for the first item

### Requirement: Validate Template

The system SHALL validate templates via `xlex template validate <template>`.

#### Scenario: Valid template
- **GIVEN** the command `xlex template validate template.xlsx`
- **WHEN** all placeholders are syntactically correct
- **THEN** exit code SHALL be 0
- **AND** output SHALL indicate "Template valid"

#### Scenario: Invalid syntax
- **GIVEN** a template with `{{#each items}` (missing closing)
- **WHEN** validated
- **THEN** exit code SHALL be non-zero
- **AND** error SHALL indicate the syntax issue and location

#### Scenario: List placeholders
- **GIVEN** the command `xlex template validate template.xlsx --list`
- **WHEN** executed
- **THEN** output SHALL list all placeholders found

#### Scenario: Generate schema
- **GIVEN** the command `xlex template validate template.xlsx --schema`
- **WHEN** executed
- **THEN** output SHALL be a JSON schema of expected data structure

### Requirement: Create Template

The system SHALL help create templates via `xlex template init`.

#### Scenario: Create from existing
- **GIVEN** the command `xlex template init report.xlsx --output template.xlsx`
- **WHEN** executed
- **THEN** a copy SHALL be created with sample placeholders added

#### Scenario: Create blank
- **GIVEN** the command `xlex template init template.xlsx --blank`
- **WHEN** executed
- **THEN** a new blank template SHALL be created with example structure

### Requirement: Template Preview

The system SHALL preview template rendering via `xlex template preview`.

#### Scenario: Preview with sample data
- **GIVEN** the command `xlex template preview template.xlsx data.json`
- **WHEN** executed
- **THEN** output SHALL show what values would be placed where
- **AND** no file SHALL be modified

#### Scenario: Preview to stdout
- **GIVEN** the command `xlex template preview template.xlsx data.json --format csv`
- **WHEN** executed
- **THEN** rendered data SHALL be output as CSV to stdout

### Requirement: Batch Template Processing

The system SHALL support batch template processing.

#### Scenario: One file per record
- **GIVEN** the command `xlex template apply template.xlsx data.json --output-dir ./reports/ --per-record`
- **WHEN** data.json contains array of records
- **THEN** one xlsx file SHALL be created per record

#### Scenario: Naming pattern
- **GIVEN** the command with `--output-pattern "report-{{id}}.xlsx"`
- **WHEN** each record has an "id" field
- **THEN** files SHALL be named using the pattern

#### Scenario: Batch from CSV
- **GIVEN** the command `xlex template apply template.xlsx records.csv --output-dir ./out/ --per-record`
- **WHEN** executed
- **THEN** one xlsx SHALL be created per CSV row

### Requirement: Template Markers

The system SHALL support special template markers for advanced features.

#### Scenario: Row repeat marker
- **GIVEN** a cell with `{{#row-repeat items}}`
- **WHEN** applying data
- **THEN** the entire row SHALL be repeated for each item

#### Scenario: Sheet repeat marker
- **GIVEN** a sheet named `{{#sheet-repeat regions}}`
- **WHEN** applying data with `{"regions": ["North", "South"]}`
- **THEN** the sheet SHALL be duplicated for each region

#### Scenario: Image placeholder
- **GIVEN** a cell with `{{#image logo}}`
- **WHEN** data contains `{"logo": "path/to/image.png"}`
- **THEN** the image SHALL be inserted at that cell

#### Scenario: Chart data marker
- **GIVEN** a chart with data range marked `{{#chart-data sales}}`
- **WHEN** applying data
- **THEN** chart data range SHALL be updated to match data length
