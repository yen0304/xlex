# Design: XLEX CLI Excel Engine

## Context

XLEX is designed as a CLI-first Excel manipulation engine for developers and automation pipelines. The primary users are:
- DevOps engineers automating report generation
- Data engineers building ETL pipelines
- Developers integrating Excel operations into scripts
- CI/CD systems processing Excel-based configurations

### Constraints

- Must handle files up to 200MB without memory exhaustion
- Must provide sub-second response for common operations
- Must be distributable as a single static binary
- Must not execute any code from xlsx files (security)

## Goals / Non-Goals

### Goals

- **Streaming architecture**: Process xlsx files without full materialization
- **CLI-first design**: Every operation accessible via command line
- **Pipeline integration**: stdin/stdout support for Unix philosophy
- **Deterministic behavior**: Same input always produces same output
- **Comprehensive operations**: Full Excel structure manipulation

### Non-Goals

- GUI or TUI interface
- Formula evaluation engine
- VBA/macro execution
- Pixel-perfect style rendering
- Real-time collaboration features

## Architecture Decisions

### Decision 0: Workspace Crate Structure

**Choice**: Organize the project as a Cargo workspace with separate crates for core library and CLI.

**Structure**:
```
xlex/
├── Cargo.toml              # Workspace root
├── crates/
│   ├── xlex-core/          # Core engine library
│   │   ├── Cargo.toml
│   │   └── src/
│   │       ├── lib.rs
│   │       ├── workbook.rs
│   │       ├── sheet.rs
│   │       ├── cell.rs
│   │       ├── parser/
│   │       ├── writer/
│   │       └── error.rs
│   └── xlex-cli/           # CLI binary
│       ├── Cargo.toml
│       └── src/
│           ├── main.rs
│           ├── commands/
│           └── output.rs
```

**Rationale**:
- **Reusability**: `xlex-core` can be used as a library by other Rust projects
- **Future bindings**: Enables PyO3 (Python), WASM, or other language bindings
- **Separation of concerns**: Core logic independent of CLI presentation
- **Testing**: Core can be tested independently of CLI

**Crate Responsibilities**:

`xlex-core`:
- ZIP/XML streaming parser
- Workbook/Sheet/Cell data structures
- Read/write operations
- SharedStrings and Style management
- Error types
- No CLI dependencies (no clap, no colored output)

`xlex-cli`:
- Command-line argument parsing (clap)
- Output formatting (text, JSON, CSV)
- Progress indicators
- Color handling
- Configuration file loading
- Calls into `xlex-core` for all operations

**Public API Design**:
```rust
// xlex-core public API
pub struct Workbook { ... }
impl Workbook {
    pub fn open(path: &Path) -> Result<Self, XlexError>;
    pub fn sheets(&self) -> impl Iterator<Item = &Sheet>;
    pub fn get_cell(&self, sheet: &str, ref_: &CellRef) -> Result<Cell, XlexError>;
    pub fn set_cell(&mut self, sheet: &str, ref_: &CellRef, value: CellValue) -> Result<(), XlexError>;
    pub fn save(&self) -> Result<(), XlexError>;
    pub fn save_as(&self, path: &Path) -> Result<(), XlexError>;
}
```

### Decision 1: SAX-based XML Parsing

**Choice**: Use `quick-xml` with event-based (SAX) parsing exclusively.

**Rationale**: DOM parsing would require loading entire XML documents into memory. For a 200MB xlsx with millions of cells, this is prohibitive. SAX parsing allows streaming through data with constant memory overhead.

**Alternatives Considered**:
- `roxmltree` (DOM): Rejected due to memory requirements
- `xml-rs` (SAX): Rejected due to performance (quick-xml is 10x faster)

### Decision 2: Lazy SharedStrings with LRU Cache

**Choice**: Load SharedStrings index on-demand with LRU cache for recently accessed strings.

**Rationale**: SharedStrings.xml can contain millions of unique strings. Loading all into memory defeats streaming. LRU cache provides O(1) access for hot strings while bounding memory.

**Cache Configuration**:
- Default: 10,000 entries
- Configurable via `--string-cache-size`

### Decision 3: Copy-on-Write ZIP Modification

**Choice**: When modifying xlsx, copy unchanged entries directly and only rewrite modified entries.

**Rationale**: Rewriting entire ZIP for single cell change is wasteful. Copy-on-write minimizes I/O and preserves entries we don't understand (future Excel features).

**Implementation**:
1. Open source ZIP for reading
2. Create new ZIP for writing
3. For each entry: copy if unchanged, rewrite if modified
4. Atomic rename on completion

### Decision 4: Internal Operation Model

**Choice**: CLI commands translate to internal operation structs, but this IR is not exposed publicly.

**Rationale**: Internal structure enables future optimizations (batching, reordering) without API commitment. Users interact only via CLI.

**Structure**:
```rust
enum Operation {
    SheetAdd { name: String },
    CellSet { sheet: String, ref_: CellRef, value: CellValue },
    // ... etc
}
```

### Decision 5: Error Handling Strategy

**Choice**: Typed errors with machine-readable codes and human messages.

**Rationale**: Scripts need parseable errors; humans need understandable messages. Both needs served by structured error type.

**Format**:
```
XLEX_E001: Sheet not found: "Summary"
```

## Component Architecture

```
┌─────────────────────────────────────────────────────────────┐
│                        CLI Layer                             │
│  ┌─────────┐ ┌─────────┐ ┌─────────┐ ┌─────────┐           │
│  │ workbook│ │  sheet  │ │  cell   │ │  range  │  ...      │
│  │ commands│ │ commands│ │ commands│ │ commands│           │
│  └────┬────┘ └────┬────┘ └────┬────┘ └────┬────┘           │
└───────┼───────────┼───────────┼───────────┼─────────────────┘
        │           │           │           │
        └───────────┴───────────┴───────────┘
                          │
┌─────────────────────────▼───────────────────────────────────┐
│                   Operation Executor                         │
│  - Validates operations                                      │
│  - Sequences dependent operations                            │
│  - Manages transaction boundaries                            │
└─────────────────────────┬───────────────────────────────────┘
                          │
┌─────────────────────────▼───────────────────────────────────┐
│                    XLSX Core Engine                          │
│  ┌──────────────┐  ┌──────────────┐  ┌──────────────┐       │
│  │ ZIP Streamer │  │  XML Parser  │  │ XML Writer   │       │
│  │ (zip crate)  │  │ (quick-xml)  │  │ (quick-xml)  │       │
│  └──────┬───────┘  └──────┬───────┘  └──────┬───────┘       │
│         │                 │                 │               │
│  ┌──────▼─────────────────▼─────────────────▼───────┐       │
│  │              Workbook State Manager               │       │
│  │  - Sheet registry                                 │       │
│  │  - SharedStrings (lazy + LRU)                    │       │
│  │  - Style registry                                 │       │
│  │  - Relationship tracker                           │       │
│  └──────────────────────────────────────────────────┘       │
└─────────────────────────────────────────────────────────────┘
                          │
┌─────────────────────────▼───────────────────────────────────┐
│                    Output Layer                              │
│  ┌────────┐  ┌────────┐  ┌────────┐  ┌────────┐            │
│  │  XLSX  │  │  CSV   │  │  JSON  │  │ NDJSON │            │
│  └────────┘  └────────┘  └────────┘  └────────┘            │
└─────────────────────────────────────────────────────────────┘
```

## Data Flow

### Read Operation (e.g., `xlex cell get`)

```
1. Open ZIP archive (memory-mapped if possible)
2. Locate sheet XML entry from workbook.xml.rels
3. Create SAX parser for sheet XML
4. Stream through rows until target row found
5. Stream through cells until target cell found
6. If cell type is 's', lookup SharedStrings (lazy load if needed)
7. Format and output value
8. Close resources
```

### Write Operation (e.g., `xlex cell set`)

```
1. Open source ZIP for reading
2. Create temp ZIP for writing
3. Copy [Content_Types].xml, _rels, docProps unchanged
4. Parse workbook.xml, copy unchanged
5. For target sheet:
   a. Stream through source XML
   b. When target cell location reached, write new value
   c. Continue streaming remainder
6. Update SharedStrings if new string added
7. Copy remaining entries unchanged
8. Atomic rename temp to target
```

## Risks / Trade-offs

### Risk: SharedStrings Index Corruption

**Scenario**: If SharedStrings.xml is malformed, string lookups fail.

**Mitigation**: 
- Validate SharedStrings structure on first access
- Fall back to inline strings if SharedStrings unavailable
- Clear error message identifying corruption

### Risk: Large File Memory Pressure

**Scenario**: 200MB file with pathological structure (e.g., one cell per row across millions of rows).

**Mitigation**:
- Strict streaming with no row buffering
- Configurable string cache size
- Memory monitoring with graceful degradation

### Trade-off: No Formula Evaluation

**Impact**: Users cannot get computed values, only formula text.

**Rationale**: Formula evaluation requires implementing Excel's entire calculation engine. Out of scope. Users should open in Excel for computation.

### Trade-off: Style Limitations

**Impact**: Cannot create arbitrary styles, only apply from predefined set or existing styles.

**Rationale**: Excel's style system is complex (fonts, fills, borders, number formats, all indexed). Full support requires significant complexity. MVP focuses on common patterns.

## Migration Plan

Not applicable - greenfield project.

## Open Questions

1. **Concurrent file access**: Should we support advisory locking for multi-process scenarios?
   - Tentative: No, keep simple. Document that concurrent writes are undefined.

2. **Undo/history**: Should modifications be reversible?
   - Tentative: No, users should maintain their own backups.

3. **Plugin system**: How should extensibility work?
   - Tentative: Defer to post-MVP. Document hook points but don't implement.

4. **XLSM support**: When to add macro-enabled workbook support?
   - Tentative: Post-MVP. Structure allows it but security review needed.
