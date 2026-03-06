# Style Operations Specification

## ADDED Requirements

### Requirement: List Styles

The system SHALL list available styles via `xlex style list <file>`.

#### Scenario: List all styles
- **GIVEN** the command `xlex style list report.xlsx`
- **WHEN** executed
- **THEN** output SHALL list all style IDs with descriptions

#### Scenario: List with details
- **GIVEN** the command `xlex style list report.xlsx --long`
- **WHEN** executed
- **THEN** output SHALL include for each style:
  - Style ID
  - Font (name, size, bold, italic, color)
  - Fill (pattern, color)
  - Border (style, color)
  - Number format

#### Scenario: JSON output
- **GIVEN** the command `xlex style list report.xlsx --format json`
- **WHEN** executed
- **THEN** output SHALL be JSON array of style definitions

### Requirement: Get Cell Style

The system SHALL retrieve cell styles via `xlex style get <file> <sheet> <ref>`.

#### Scenario: Get style details
- **GIVEN** the command `xlex style get report.xlsx "Data" A1`
- **WHEN** executed
- **THEN** output SHALL include all style properties of A1

#### Scenario: Get style ID only
- **GIVEN** the command `xlex style get report.xlsx "Data" A1 --id-only`
- **WHEN** executed
- **THEN** output SHALL be only the style ID number

#### Scenario: JSON output
- **GIVEN** the command `xlex style get report.xlsx "Data" A1 --format json`
- **WHEN** executed
- **THEN** output SHALL be JSON object with all style properties

### Requirement: Apply Style to Range

The system SHALL apply styles via `xlex range style <file> <sheet> <range>`.

#### Scenario: Apply bold
- **GIVEN** the command `xlex range style report.xlsx "Data" A1:D1 --bold`
- **WHEN** executed
- **THEN** all cells in A1:D1 SHALL have bold font

#### Scenario: Apply italic
- **GIVEN** the command `xlex range style report.xlsx "Data" A1:D1 --italic`
- **WHEN** executed
- **THEN** all cells in A1:D1 SHALL have italic font

#### Scenario: Apply underline
- **GIVEN** the command `xlex range style report.xlsx "Data" A1:D1 --underline`
- **WHEN** executed
- **THEN** all cells in A1:D1 SHALL have underlined text

#### Scenario: Apply font size
- **GIVEN** the command `xlex range style report.xlsx "Data" A1:D1 --font-size 14`
- **WHEN** executed
- **THEN** all cells in A1:D1 SHALL have font size 14

#### Scenario: Apply font name
- **GIVEN** the command `xlex range style report.xlsx "Data" A1:D1 --font "Arial"`
- **WHEN** executed
- **THEN** all cells in A1:D1 SHALL use Arial font

#### Scenario: Apply font color
- **GIVEN** the command `xlex range style report.xlsx "Data" A1:D1 --color "#FF0000"`
- **WHEN** executed
- **THEN** all cells in A1:D1 SHALL have red text

#### Scenario: Apply background color
- **GIVEN** the command `xlex range style report.xlsx "Data" A1:D1 --bg-color "#FFFF00"`
- **WHEN** executed
- **THEN** all cells in A1:D1 SHALL have yellow background

#### Scenario: Apply horizontal alignment
- **GIVEN** the command `xlex range style report.xlsx "Data" A1:D1 --align center`
- **WHEN** executed
- **THEN** all cells in A1:D1 SHALL be horizontally centered

#### Scenario: Apply vertical alignment
- **GIVEN** the command `xlex range style report.xlsx "Data" A1:D1 --valign middle`
- **WHEN** executed
- **THEN** all cells in A1:D1 SHALL be vertically centered

#### Scenario: Apply text wrap
- **GIVEN** the command `xlex range style report.xlsx "Data" A1:D1 --wrap`
- **WHEN** executed
- **THEN** all cells in A1:D1 SHALL have text wrapping enabled

#### Scenario: Apply number format
- **GIVEN** the command `xlex range style report.xlsx "Data" B2:B100 --number-format "#,##0.00"`
- **WHEN** executed
- **THEN** all cells SHALL display with thousands separator and 2 decimals

#### Scenario: Apply date format
- **GIVEN** the command `xlex range style report.xlsx "Data" C2:C100 --date-format "YYYY-MM-DD"`
- **WHEN** executed
- **THEN** all cells SHALL display dates in ISO format

#### Scenario: Apply percentage format
- **GIVEN** the command `xlex range style report.xlsx "Data" D2:D100 --percent`
- **WHEN** executed
- **THEN** all cells SHALL display as percentages

#### Scenario: Apply currency format
- **GIVEN** the command `xlex range style report.xlsx "Data" E2:E100 --currency USD`
- **WHEN** executed
- **THEN** all cells SHALL display with $ symbol

#### Scenario: Apply multiple styles
- **GIVEN** the command `xlex range style report.xlsx "Data" A1:D1 --bold --bg-color "#4472C4" --color "#FFFFFF" --align center`
- **WHEN** executed
- **THEN** all specified styles SHALL be applied

### Requirement: Apply Borders

The system SHALL apply borders via `xlex range border <file> <sheet> <range>`.

#### Scenario: Apply all borders
- **GIVEN** the command `xlex range border report.xlsx "Data" A1:D10 --all`
- **WHEN** executed
- **THEN** all cells SHALL have borders on all sides

#### Scenario: Apply outline border
- **GIVEN** the command `xlex range border report.xlsx "Data" A1:D10 --outline`
- **WHEN** executed
- **THEN** only the outer edge of the range SHALL have borders

#### Scenario: Apply specific borders
- **GIVEN** the command `xlex range border report.xlsx "Data" A1:D1 --bottom`
- **WHEN** executed
- **THEN** only bottom borders SHALL be applied

#### Scenario: Apply border style
- **GIVEN** the command `xlex range border report.xlsx "Data" A1:D10 --all --style thick`
- **WHEN** executed
- **THEN** borders SHALL be thick style

#### Scenario: Apply border color
- **GIVEN** the command `xlex range border report.xlsx "Data" A1:D10 --all --border-color "#000000"`
- **WHEN** executed
- **THEN** borders SHALL be black

#### Scenario: Remove borders
- **GIVEN** the command `xlex range border report.xlsx "Data" A1:D10 --none`
- **WHEN** executed
- **THEN** all borders SHALL be removed from the range

### Requirement: Style Presets

The system SHALL support style presets via `xlex style preset`.

#### Scenario: List presets
- **GIVEN** the command `xlex style preset list`
- **WHEN** executed
- **THEN** output SHALL list built-in presets:
  - header (bold, centered, background)
  - currency (number format, right-aligned)
  - percent (percentage format)
  - date (date format)
  - error (red text)
  - warning (yellow background)

#### Scenario: Apply preset
- **GIVEN** the command `xlex style preset apply report.xlsx "Data" A1:D1 header`
- **WHEN** executed
- **THEN** the header preset styles SHALL be applied

#### Scenario: Create custom preset
- **GIVEN** the command `xlex style preset create "myheader" --bold --bg-color "#4472C4" --color "#FFFFFF"`
- **WHEN** executed
- **THEN** a custom preset "myheader" SHALL be saved

#### Scenario: Delete custom preset
- **GIVEN** the command `xlex style preset delete "myheader"`
- **WHEN** executed
- **THEN** the custom preset SHALL be removed

### Requirement: Copy Style

The system SHALL copy styles via `xlex style copy <file> <sheet> <source> <dest>`.

#### Scenario: Copy cell style
- **GIVEN** the command `xlex style copy report.xlsx "Data" A1 B1:D1`
- **WHEN** executed
- **THEN** A1's style SHALL be applied to B1:D1
- **AND** values SHALL not be affected

#### Scenario: Copy to another sheet
- **GIVEN** the command `xlex style copy report.xlsx "Data" A1:D1 --to-sheet "Summary" A1:D1`
- **WHEN** executed
- **THEN** styles SHALL be copied to Summary sheet

### Requirement: Clear Style

The system SHALL clear styles via `xlex style clear <file> <sheet> <range>`.

#### Scenario: Clear all styles
- **GIVEN** the command `xlex style clear report.xlsx "Data" A1:D10`
- **WHEN** executed
- **THEN** all cells SHALL revert to default style
- **AND** values SHALL be preserved

#### Scenario: Clear specific style
- **GIVEN** the command `xlex style clear report.xlsx "Data" A1:D10 --font-only`
- **WHEN** executed
- **THEN** only font styles SHALL be cleared

### Requirement: Conditional Formatting Setup

The system SHALL set up conditional formatting rules via `xlex style condition`.

#### Scenario: Add highlight rule
- **GIVEN** the command `xlex style condition report.xlsx "Data" B2:B100 --highlight-cells --gt 1000 --bg-color "#90EE90"`
- **WHEN** executed
- **THEN** cells > 1000 SHALL have green background when opened in Excel

#### Scenario: Add color scale
- **GIVEN** the command `xlex style condition report.xlsx "Data" B2:B100 --color-scale --min "#FF0000" --max "#00FF00"`
- **WHEN** executed
- **THEN** a red-to-green color scale SHALL be applied

#### Scenario: Add data bars
- **GIVEN** the command `xlex style condition report.xlsx "Data" B2:B100 --data-bars --color "#4472C4"`
- **WHEN** executed
- **THEN** data bars SHALL be added to the range

#### Scenario: Add icon set
- **GIVEN** the command `xlex style condition report.xlsx "Data" B2:B100 --icon-set "3-arrows"`
- **WHEN** executed
- **THEN** icon set conditional formatting SHALL be applied

#### Scenario: List conditional formats
- **GIVEN** the command `xlex style condition report.xlsx "Data" --list`
- **WHEN** executed
- **THEN** output SHALL list all conditional formatting rules

#### Scenario: Remove conditional format
- **GIVEN** the command `xlex style condition report.xlsx "Data" B2:B100 --remove`
- **WHEN** executed
- **THEN** conditional formatting SHALL be removed from the range

### Requirement: Freeze Panes

The system SHALL manage freeze panes via `xlex style freeze <file> <sheet>`.

#### Scenario: Freeze rows
- **GIVEN** the command `xlex style freeze report.xlsx "Data" --rows 1`
- **WHEN** executed
- **THEN** row 1 SHALL be frozen (header row)

#### Scenario: Freeze columns
- **GIVEN** the command `xlex style freeze report.xlsx "Data" --cols 1`
- **WHEN** executed
- **THEN** column A SHALL be frozen

#### Scenario: Freeze at cell
- **GIVEN** the command `xlex style freeze report.xlsx "Data" --at B2`
- **WHEN** executed
- **THEN** rows above and columns left of B2 SHALL be frozen

#### Scenario: Unfreeze
- **GIVEN** the command `xlex style freeze report.xlsx "Data" --unfreeze`
- **WHEN** executed
- **THEN** all freeze panes SHALL be removed

#### Scenario: Get freeze status
- **GIVEN** the command `xlex style freeze report.xlsx "Data"`
- **WHEN** executed without flags
- **THEN** output SHALL show current freeze pane settings
