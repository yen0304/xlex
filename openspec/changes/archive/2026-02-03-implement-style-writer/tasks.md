# Tasks: Implement Style Writer

## Implementation Checklist

### Phase 1: Preparation
- [x] **1.1** Review OOXML styles.xml specification and Excel-generated samples
- [x] **1.2** Verify `Style`, `Font`, `Fill`, `Border` structs have `PartialEq` for deduplication
- [x] **1.3** Add helper methods to style types for OOXML output (e.g., `Color::to_argb_hex()`)

### Phase 2: Core Writer Implementation  
- [x] **2.1** Create `StyleCollector` struct to gather unique fonts/fills/borders/numfmts from `StyleRegistry`
- [x] **2.2** Implement `write_fonts()` - generate `<fonts>` section from collected fonts
- [x] **2.3** Implement `write_fills()` - generate `<fills>` section (ensure none/gray125 defaults)
- [x] **2.4** Implement `write_borders()` - generate `<borders>` section
- [x] **2.5** Implement `write_number_formats()` - generate `<numFmts>` for custom formats
- [x] **2.6** Implement `write_cell_xfs()` - generate `<cellXfs>` with proper index references
- [x] **2.7** Update `write_styles()` to use new dynamic generation

### Phase 3: Style ID Mapping
- [x] **3.1** Build mapping from `StyleRegistry` style IDs to `cellXfs` indices
- [x] **3.2** Update `write_sheet()` to use mapped indices in cell `s` attributes
- [x] **3.3** Handle case where cells reference styles not in registry (use default)

### Phase 4: Testing
- [x] **4.1** Add unit test: empty StyleRegistry produces valid minimal styles.xml
- [x] **4.2** Add unit test: single bold style produces correct fonts/cellXfs
- [x] **4.3** Add unit test: background color produces correct fills/cellXfs  
- [x] **4.4** Add unit test: border style produces correct borders/cellXfs
- [x] **4.5** Add integration test: `range style --bg-color` persists to file
- [x] **4.6** Add integration test: `range style --bold --text-color` persists to file
- [x] **4.7** Add round-trip test: open styled xlsx → save → verify styles preserved

### Phase 4.5: Parser Enhancement (Added)
- [x] **4.5.1** Update StylesParser to parse `cellXfs` elements
- [x] **4.5.2** Build complete `Style` objects from fontId/fillId/borderId/numFmtId references  
- [x] **4.5.3** Add `StyleRegistry::add_with_id()` method for loading styles with specific IDs

### Phase 5: Validation
- [x] **5.1** Run all existing tests (`cargo test`) - 622 passed, 0 failed
- [x] **5.2** Manual test: create styled xlsx, open in Excel, verify appearance
- [x] **5.3** Manual test: create styled xlsx, open in LibreOffice Calc, verify appearance
- [x] **5.4** Run clippy and fix any warnings

## Dependencies
- Task 2.* depends on 1.*
- Task 3.* depends on 2.*
- Task 4.* can partially parallelize with 3.*
- Task 5.* requires all previous tasks

## Notes
- OOXML requires fills[0]=none, fills[1]=gray125 - don't remove these defaults
- Color format is ARGB (e.g., `FFFF0000` for red, FF prefix is alpha)
- Custom number format IDs must be >= 164 (0-163 are built-in)
