# Change: Add XLEX CLI Excel Engine

## Why

There is no CLI-first, streaming-based Excel manipulation tool that treats xlsx files as structured documents. Existing solutions either require GUI interaction, load entire files into memory, or lack comprehensive Excel operation support. XLEX fills this gap by providing a high-performance, scriptable, pipeline-friendly tool for Excel automation.

## What Changes

This is a greenfield implementation establishing the complete XLEX CLI tool:

### Architecture
- **Workspace Structure**: Cargo workspace with `xlex-core` (library) and `xlex-cli` (binary)
- **Reusable Core**: `xlex-core` can be used as a Rust library for other projects
- **Future-proof**: Architecture supports future Python bindings, WASM, etc.

### Core Capabilities
- **Core Engine**: Streaming ZIP/XML parser with lazy SharedStrings and Style registry
- **Workbook Operations**: info, validate, clone, create, props, stats
- **Sheet Operations**: list, add, remove, rename, copy, move, hide/unhide, info, active
- **Cell Operations**: get, set, formula, clear, type, batch, comment, link
- **Row Operations**: get, append, insert, delete, copy, move, height, hide/unhide, find
- **Column Operations**: get, insert, delete, copy, move, width, hide/unhide, header, find, stats
- **Range Operations**: get, copy, move, clear, fill, merge/unmerge, name, validate, sort, filter
- **Style Operations**: list, get, apply, border, preset, copy, clear, condition, freeze
- **Import/Export**: CSV, JSON, NDJSON with streaming support, convert command

### New Features (Added)
- **Formula Operations**: validate, list, stats, refs, replace - for CI/CD validation and analysis
- **Template Operations**: apply, validate, init, preview - Handlebars-like template engine for report generation
- **Project Config**: `.xlex.yml` support for team-wide configuration

### CLI Interface
- Unified command structure, global flags, completion, config, batch, alias, interactive
- Project-level configuration file (`.xlex.yml` / `.xlex.yaml`)
- Config precedence: CLI > Environment > Project Config > User Config > Defaults
- Error Handling: Machine-readable error codes (XLEX_E001-E099) with human-friendly messages

### GitHub Open Source Infrastructure
- **Community Files**: README.md, LICENSE (MIT), CODE_OF_CONDUCT.md, CONTRIBUTING.md, SECURITY.md, CHANGELOG.md
- **Issue Templates**: Bug report, Feature request, Question (YAML format)
- **PR Template**: Checklist for tests, docs, changelog
- **CI/CD Workflows**: Test matrix (Linux/macOS/Windows), Release automation, Dependabot
- **Distribution**: GitHub Releases, curl install, Homebrew, cargo, npx

### Comprehensive Documentation
- **Getting Started**: Installation (all methods), Quick start, First steps
- **Command Reference**: All 90+ commands with synopsis, options, examples
- **Guides**: Pipeline integration, Automation, Large files, Error handling, Templates
- **Cookbook**: Common tasks, Data migration, Reporting, Template recipes
- **Reference**: CLI reference, Error codes, Exit codes, Environment variables, Config file, Template syntax
- **Development**: Architecture, Contributing, Building, Library usage

## Impact

- **Affected specs**: 15 capability modules (see specs/ directory)
- **Affected code**: Entire codebase (greenfield)
- **Breaking changes**: None (new project)
- **Dependencies**: zip, quick-xml, clap, serde, serde_json, serde_yaml

## Success Criteria

1. All CLI commands functional with documented behavior
2. Performance guarantees met (see spec)
3. Streaming architecture verified with 200MB test files
4. Cross-platform builds (Linux, macOS, Windows)
5. Distribution channels configured (cargo, brew, npx, curl)
6. GitHub Community Standards 100% compliance
7. Documentation site deployed with all commands documented
8. Issue templates and CI/CD workflows operational
9. `xlex-core` publishable as independent library on crates.io
10. Template engine supports common report generation use cases
