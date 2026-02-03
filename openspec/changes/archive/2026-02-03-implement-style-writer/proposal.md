# Proposal: Implement Style Writer

## Change ID
`implement-style-writer`

## Summary
Implement proper styles.xml output in the workbook writer so that cell formatting (colors, fonts, borders, etc.) applied via CLI commands actually persists to the saved xlsx file.

## Problem Statement
Currently, the xlex CLI has style-related commands (`range style`, `range border`) that:
1. Create `Style` objects with fonts, fills, borders, number formats
2. Register styles in `StyleRegistry` and assign style IDs to cells
3. Save the workbook

However, the **workbook writer outputs a hardcoded minimal styles.xml** that ignores the actual `StyleRegistry` content. This means:
- `xlex range style ... --bg-color FFFF00 --bold` reports success but the styling is lost on save
- Users expect styling to persist, but Excel shows no formatting

## Root Cause
In `crates/xlex-core/src/writer/workbook.rs`, the `write_styles` function outputs a static template:

```rust
fn write_styles<W: Write + std::io::Seek>(
    &self,
    zip: &mut ZipWriter<W>,
    _workbook: &Workbook,  // <-- workbook is ignored!
    options: SimpleFileOptions,
) -> XlexResult<()> {
    // Minimal styles.xml - hardcoded, ignores workbook.style_registry()
}
```

## Proposed Solution
Modify `write_styles` to:
1. Read fonts, fills, borders, number formats from `workbook.style_registry()`
2. Generate proper OOXML `<styleSheet>` elements dynamically
3. Map style IDs to `<cellXfs>` entries
4. Preserve existing styles when editing files (round-trip support)

## Impact
- **xlex-core**: Modify `writer/workbook.rs` to implement dynamic style output
- **xlex-core**: May need to enhance `StyleRegistry` for better ID tracking
- **No CLI changes needed**: Existing commands will work once writer is fixed

## Success Criteria
1. `xlex range style file.xlsx Sheet A1:B2 --bg-color FFFF00` produces a file where Excel shows yellow background
2. `xlex range style file.xlsx Sheet A1 --bold --text-color FF0000` produces red bold text
3. Existing styled xlsx files opened and saved retain their styles
4. All existing tests pass

## Alternatives Considered

### Alternative 1: Use SharedStrings for rich text
Not applicable - styles are cell-level, not text-level.

### Alternative 2: Write styles.xml as-is from original file
Partially solves round-trip but doesn't help new styles. Need both preservation AND new style generation.

## Risks
- **OOXML complexity**: styles.xml has interdependent sections (fonts → fills → borders → cellXfs). Must maintain proper indices.
- **Round-trip fidelity**: When editing existing files, must preserve original styles and only add new ones.

## Timeline Estimate
- Design: 1 hour
- Implementation: 4-6 hours
- Testing: 2 hours
