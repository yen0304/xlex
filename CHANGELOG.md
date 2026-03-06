# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

## [0.3.0] - 2026-03-06

### Removed

- **MCP Server**: Removed `xlex-mcp` crate entirely — the MCP approach is replaced by agent skill files
- **tokio dependency**: No longer needed without MCP server

### Added

- **Agent Skill**: New `docs/skills/xlex-agent/` with progressive disclosure structure
  - `SKILL.md` — core workflows, loaded when skill triggers
  - `references/commands.md` — complete CLI command reference (123 commands)
  - `references/examples.md` — 8 real-world workflow examples
- **Documentation coverage**: All 123 CLI commands now fully documented with correct flags

### Fixed

- **Clippy**: Resolved all clippy warnings (approx_constant, useless format!, is_ok+unwrap, default on unit struct)
- **Skill accuracy**: Fixed 6 incorrect flag names, added 15 undocumented flags, added 6 missing commands

### Changed

- **CI**: Removed `msrv-mcp` job; simplified MSRV to single 1.80 check
- **Release**: Removed xlex-mcp from packaging pipeline
- **README**: MCP Server section replaced with AI Agent Integration

## [0.2.4] - 2026-02-03

### Fixed

- **Range Style**: Fixed `range style --bold` command now correctly applies styles to cells
- **Integration Tests**: Updated range style test to verify style application

### Improved

- **SharedStrings Deduplication**: Added O(1) HashMap lookup for shared string deduplication
  - Previously O(n) linear search, now instant lookup
  - Significant performance improvement for workbooks with many repeated strings

## [0.2.0] - 2026-02-03

### Added

- **Session Mode**: `xlex session <file>` command for efficient large file operations
  - Load file once, run multiple commands instantly
  - Supported commands: info, sheets, cell, row, help, exit
- **LazyWorkbook API**: New lazy-loading workbook type for large files
  - `LazyWorkbook::open()` - metadata-only loading
  - `stream_rows()` - streaming row iterator
  - `read_cell()` - single cell access without full load
- **Parallel Sheet Parsing**: Optional `parallel` feature using rayon

### Changed

- **Memory-mapped I/O**: Files >10MB automatically use mmap for better performance
- **Lazy SharedStrings**: On-demand parsing with LRU cache (default 10,000 entries)

### Performance

- 366MB file load time: **0.23s** (was 86s) - **190x faster**
- Session mode enables instant subsequent commands after initial load

## [0.1.0] - 2026-02-02

### Added

- Initial release of xlex CLI
- **Core Engine**: Streaming ZIP/XML parser with lazy SharedStrings
- **Workbook Operations**: info, validate, clone, create, props, stats
- **Sheet Operations**: list, add, remove, rename, copy, move, hide/unhide, info, active
- **Cell Operations**: get, set, formula, clear, type, batch, comment, link
- **Row Operations**: get, append, insert, delete, copy, move, height, hide/unhide, find
- **Column Operations**: get, insert, delete, copy, move, width, hide/unhide, header, find, stats
- **Range Operations**: get, copy, move, clear, fill, merge/unmerge, name, validate, sort, filter
- **Style Operations**: list, get, apply, border, preset, copy, clear, condition, freeze
- **Import/Export**: CSV, JSON, NDJSON with streaming support
- **Formula Operations**: validate, list, stats, refs, replace
- **Template Operations**: apply, validate, init, preview with Handlebars-like syntax
- **CLI Features**: completion, config, batch, alias, interactive mode
- Cross-platform support (Linux, macOS, Windows)
- Installation via cargo, brew, npm, curl

[Unreleased]: https://github.com/yen0304/xlex/compare/v0.3.0...HEAD
[0.3.0]: https://github.com/yen0304/xlex/compare/v0.2.4...v0.3.0
[0.2.4]: https://github.com/yen0304/xlex/compare/v0.2.0...v0.2.4
[0.2.0]: https://github.com/yen0304/xlex/compare/v0.1.0...v0.2.0
[0.1.0]: https://github.com/yen0304/xlex/releases/tag/v0.1.0
