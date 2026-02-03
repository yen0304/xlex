# Spec: Style Persistence

## Overview
Cell styling applied through xlex commands must persist when saving xlsx files. Styles include fonts, colors, fills, borders, alignments, and number formats.

## ADDED Requirements

### Requirement: Font Style Persistence
Font properties (bold, italic, underline, color, size, name) applied to cells MUST be written to styles.xml and MUST appear correctly when the file is opened in Excel or other spreadsheet applications.

#### Scenario: Bold text persists after save
**Given** a workbook with a cell containing text  
**When** bold style is applied via `range style --bold`  
**And** the workbook is saved  
**Then** opening the file in Excel shows the cell text as bold

#### Scenario: Text color persists after save
**Given** a workbook with a cell containing text  
**When** text color is applied via `range style --text-color FF0000`  
**And** the workbook is saved  
**Then** opening the file in Excel shows red text

#### Scenario: Multiple font properties combine correctly
**Given** a workbook with a cell containing text  
**When** `range style --bold --italic --text-color 0000FF --font-size 14` is applied  
**And** the workbook is saved  
**Then** opening the file shows bold italic blue 14pt text

---

### Requirement: Fill Style Persistence
Background fill patterns and colors applied to cells MUST be written to styles.xml and MUST display correctly.

#### Scenario: Background color persists after save
**Given** a workbook with a cell  
**When** background color is applied via `range style --bg-color FFFF00`  
**And** the workbook is saved  
**Then** opening the file in Excel shows yellow cell background

#### Scenario: Background color on multiple cells
**Given** a workbook with cells A1:C3  
**When** `range style A1:C3 --bg-color 00FF00` is applied  
**And** the workbook is saved  
**Then** all 9 cells show green background in Excel

---

### Requirement: Border Style Persistence
Cell borders applied through commands MUST persist to the saved file.

#### Scenario: Thin border persists after save
**Given** a workbook with cells A1:B2  
**When** `range border --all --style thin` is applied  
**And** the workbook is saved  
**Then** opening the file shows thin borders around and between all cells

#### Scenario: Colored border persists after save
**Given** a workbook with a cell  
**When** `range border --all --border-color FF0000` is applied  
**And** the workbook is saved  
**Then** the cell shows red borders in Excel

---

### Requirement: Number Format Persistence
Custom number formats applied to cells MUST persist and MUST display values correctly.

#### Scenario: Percentage format persists
**Given** a workbook with cell A1 containing 0.15  
**When** `range style A1 --percent` is applied  
**And** the workbook is saved  
**Then** Excel displays "15%" in cell A1

#### Scenario: Custom number format persists
**Given** a workbook with cell A1 containing 1234.5  
**When** `range style A1 --number-format "#,##0.00"` is applied  
**And** the workbook is saved  
**Then** Excel displays "1,234.50" in cell A1

---

### Requirement: Style Round-trip Preservation
When opening an existing styled xlsx file and saving it, all original styles MUST be preserved.

#### Scenario: Existing styles preserved on save
**Given** an xlsx file with various cell styles created by Excel  
**When** the file is opened with xlex  
**And** a cell value is modified  
**And** the file is saved  
**Then** all original cell styles remain intact

#### Scenario: New styles added to existing styled file
**Given** an xlsx file with cells A1:A5 having styles created by Excel  
**When** the file is opened with xlex  
**And** `range style B1:B5 --bold` is applied  
**And** the file is saved  
**Then** original A1:A5 styles are preserved  
**And** B1:B5 shows bold text

---

### Requirement: Alignment Persistence  
Text alignment settings MUST persist when saving.

#### Scenario: Horizontal alignment persists
**Given** a workbook with a cell  
**When** `range style --align center` is applied  
**And** the workbook is saved  
**Then** Excel shows the cell content centered horizontally

#### Scenario: Text wrap persists
**Given** a workbook with a cell containing long text  
**When** `range style --wrap` is applied  
**And** the workbook is saved  
**Then** Excel shows the text wrapped within the cell

## Technical Notes
- styles.xml must maintain required OOXML structure: fonts → fills → borders → cellXfs
- fills[0] must be "none" pattern, fills[1] must be "gray125" (OOXML requirement)
- Color values in OOXML use ARGB format (e.g., FFFF0000 for opaque red)
- Custom number format IDs must be >= 164 (0-163 are reserved for built-in formats)
