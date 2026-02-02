# Project Context

## Purpose

XLEX is a CLI-first Excel manipulation engine that treats `.xlsx` files as structured documents rather than GUI spreadsheets. It provides high-performance, streaming-based operations for programmatic Excel file manipulation.

**Core Philosophy**: Excel is data + schema + rules, not UI.

## Tech Stack

- **Language**: Rust (for performance and safety)
- **ZIP Handling**: `zip` crate
- **XML Parsing**: `quick-xml` (SAX/event-based for streaming)
- **CLI Framework**: `clap`
- **Serialization**: `serde` with `serde_json` and `serde_yaml`

## Project Conventions

### Code Style

- Follow Rust standard naming conventions (snake_case for functions/variables, PascalCase for types)
- Use `rustfmt` for formatting
- Use `clippy` for linting with pedantic warnings enabled
- Prefer explicit error handling over panics
- Document all public APIs with rustdoc

### Architecture Patterns

- **Streaming-first**: Never materialize full tables in memory
- **Zero-copy where possible**: Use references and iterators
- **Copy-on-write for modifications**: Only rewrite changed ZIP entries
- **Lazy evaluation**: SharedStrings and styles loaded on-demand
- **SAX parsing only**: No DOM-based XML parsing

### Testing Strategy

- Unit tests for all core data structures
- Integration tests for CLI commands
- Property-based tests for parsers
- Benchmark tests for performance guarantees
- Test fixtures with real xlsx files of various sizes

### Git Workflow

- `main` branch is always deployable
- Feature branches: `feat/description`
- Bug fixes: `fix/description`
- Conventional commits: `feat:`, `fix:`, `docs:`, `test:`, `refactor:`

## Domain Context

### XLSX File Structure

An `.xlsx` file is a ZIP archive containing:
- `[Content_Types].xml` - MIME type declarations
- `_rels/.rels` - Package relationships
- `xl/workbook.xml` - Workbook structure and sheet references
- `xl/worksheets/sheet*.xml` - Individual sheet data
- `xl/sharedStrings.xml` - Deduplicated string table
- `xl/styles.xml` - Cell formatting definitions
- `xl/_rels/workbook.xml.rels` - Workbook relationships
- `docProps/core.xml` - Document properties

### Cell Types

- `s` - Shared string (index into sharedStrings.xml)
- `n` - Number
- `b` - Boolean
- `str` - Inline string
- `e` - Error
- `d` - Date (ISO 8601)

### Cell References

- A1 notation: Column letter(s) + row number (e.g., A1, Z100, AA1)
- R1C1 notation: Not supported (A1 only)
- Range notation: Start:End (e.g., A1:B10)

## Important Constraints

- **No VBA execution**: Security boundary - never execute macros
- **No formula evaluation**: Formulas are stored, not computed
- **No pixel-perfect rendering**: Structure over presentation
- **Memory bound**: Must handle 200MB files without OOM
- **Performance SLAs**: See performance guarantees in specs

## External Dependencies

- No external services required
- No network calls during operation
- Self-contained single binary distribution
