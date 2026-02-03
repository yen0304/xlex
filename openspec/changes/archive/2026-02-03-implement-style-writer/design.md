# Design: Style Writer Implementation

## Architecture Overview

### Current State
```
CLI Command                    Core Library                    Output
───────────────────────────────────────────────────────────────────────
range style --bg-color  ──►  StyleRegistry.add(style)   ──►  styles.xml
                              cell.style_id = N              (HARDCODED)
                                     │
                                     └──► LOST! Writer ignores StyleRegistry
```

### Target State
```
CLI Command                    Core Library                    Output
───────────────────────────────────────────────────────────────────────
range style --bg-color  ──►  StyleRegistry.add(style)   ──►  styles.xml
                              cell.style_id = N              (DYNAMIC)
                                     │
                                     └──► write_styles() reads StyleRegistry
                                          and generates proper XML
```

## OOXML styles.xml Structure

The styles.xml file has a strict structure with interdependent sections:

```xml
<styleSheet>
    <numFmts>           <!-- Custom number formats (id >= 164) -->
    <fonts>             <!-- Font definitions, indexed 0..N -->
    <fills>             <!-- Fill patterns, indexed 0..N (0=none, 1=gray125 required) -->
    <borders>           <!-- Border definitions, indexed 0..N -->
    <cellStyleXfs>      <!-- Base styles (usually just one default) -->
    <cellXfs>           <!-- Cell formats - references fonts/fills/borders by index -->
</styleSheet>
```

### Index Relationships
- `cellXfs[i]` references `fontId`, `fillId`, `borderId`, `numFmtId`
- Cell's `s` attribute in sheet XML = index into `cellXfs`
- Default cellXfs[0] should be the "Normal" style

## Implementation Strategy

### Phase 1: Build Index Maps
When writing styles.xml, collect all unique components:

```rust
struct StyleComponents {
    fonts: Vec<Font>,           // fonts[i] → fontId = i
    fills: Vec<Fill>,           // fills[i] → fillId = i
    borders: Vec<Border>,       // borders[i] → borderId = i
    num_formats: Vec<(u32, String)>, // custom formats with id >= 164
    cell_xfs: Vec<CellXf>,      // final cell format combinations
}

struct CellXf {
    font_id: usize,
    fill_id: usize,
    border_id: usize,
    num_fmt_id: u32,
    alignment: Option<Alignment>,
}
```

### Phase 2: Deduplication
Avoid duplicate entries by comparing components:

```rust
fn find_or_add_font(fonts: &mut Vec<Font>, font: &Font) -> usize {
    fonts.iter().position(|f| f == font)
        .unwrap_or_else(|| {
            fonts.push(font.clone());
            fonts.len() - 1
        })
}
```

### Phase 3: XML Generation
Generate each section in order, maintaining indices:

```rust
fn write_styles(&self, zip: &mut ZipWriter<W>, workbook: &Workbook, ...) {
    let registry = workbook.style_registry();
    let components = self.collect_style_components(registry);
    
    // Write numFmts (if custom formats exist)
    // Write fonts
    // Write fills (must start with none, gray125)
    // Write borders
    // Write cellStyleXfs
    // Write cellXfs
}
```

## Key Decisions

### Decision 1: Minimal Required Fills
OOXML requires fills[0] = none, fills[1] = gray125. Always include these.

### Decision 2: Style ID Mapping
Current `StyleRegistry` assigns arbitrary IDs. Writer must map these to contiguous `cellXfs` indices.

```rust
// Map: registry style ID → cellXfs index
style_id_to_xf_index: HashMap<u32, usize>
```

### Decision 3: Round-trip Preservation
When opening an existing file:
1. Parser already loads styles into `StyleRegistry`
2. On save, regenerate styles.xml from `StyleRegistry`
3. As long as parser preserves all style info, round-trip works

### Decision 4: Font Color Format
OOXML uses ARGB format: `<color rgb="FFFF0000"/>` (FF prefix for alpha).

## File Changes

### `crates/xlex-core/src/writer/workbook.rs`
- Replace hardcoded `write_styles` with dynamic implementation
- Add helper functions for each style section

### `crates/xlex-core/src/style.rs` (minor)
- Ensure `PartialEq` derives for deduplication
- May need `to_ooxml()` methods for XML generation

## Testing Strategy

1. **Unit tests**: Generate styles.xml for known `StyleRegistry` states
2. **Round-trip tests**: Open styled xlsx → save → reopen → verify styles
3. **Integration tests**: CLI commands produce correct Excel rendering
4. **Manual verification**: Open in Excel/LibreOffice to confirm visual output
