# performance Specification

## Purpose
TBD - created by archiving change optimize-large-file-performance. Update Purpose after archive.
## Requirements
### Requirement: Large File Handling

The system SHALL handle xlsx files up to 500MB without memory exhaustion.

#### Scenario: Open 300MB file
- **GIVEN** an xlsx file of 300MB size
- **WHEN** `Workbook::open()` is called
- **THEN** the operation completes in less than 10 seconds
- **AND** peak memory usage is less than 600MB

#### Scenario: Metadata query on large file
- **GIVEN** an xlsx file of 300MB size
- **WHEN** only sheet names are queried via `sheet_names()`
- **THEN** the full file content is NOT loaded into memory
- **AND** the operation completes in less than 2 seconds

---

### Requirement: Lazy SharedStrings Loading

The system SHALL support lazy loading of shared strings with on-demand parsing.

#### Scenario: Index-based string access
- **GIVEN** a workbook with 1 million shared strings
- **WHEN** a single string is accessed by index
- **THEN** only that string is parsed from the source
- **AND** the result is cached for subsequent access

#### Scenario: LRU cache eviction
- **GIVEN** an LRU cache with capacity N
- **WHEN** more than N unique strings are accessed
- **THEN** least recently used strings are evicted
- **AND** re-accessing evicted strings re-parses from source

---

### Requirement: Memory-Mapped File Access

The system SHALL use memory mapping for large file access when supported by the platform.

#### Scenario: Automatic mmap selection
- **GIVEN** a file larger than the mmap threshold (default 10MB)
- **WHEN** the file is opened
- **THEN** memory mapping is used instead of buffered I/O

#### Scenario: Fallback for unsupported sources
- **GIVEN** input from stdin (not a file)
- **WHEN** opening the workbook
- **THEN** buffered I/O is used as fallback
- **AND** all functionality remains available

---

### Requirement: Parallel Sheet Parsing

The system SHALL support parallel parsing of multiple sheets when the `parallel` feature is enabled.

#### Scenario: Multi-sheet workbook parsing
- **GIVEN** a workbook with 10 sheets
- **AND** the `parallel` feature is enabled
- **WHEN** the workbook is opened
- **THEN** sheets are parsed in parallel using available CPU cores
- **AND** the result is equivalent to sequential parsing

#### Scenario: Single sheet fallback
- **GIVEN** a workbook with only 1 sheet
- **WHEN** the workbook is opened
- **THEN** parallel parsing overhead is avoided
- **AND** sequential parsing is used

---

### Requirement: Streaming Row Access

The system SHALL provide an iterator-based API for streaming row access.

#### Scenario: Stream all rows
- **GIVEN** a sheet with 1 million rows
- **WHEN** `stream_rows()` is called
- **THEN** an iterator is returned immediately
- **AND** rows are parsed on-demand as the iterator advances
- **AND** memory usage remains constant regardless of total row count

#### Scenario: Stream row range
- **GIVEN** a sheet with 1 million rows
- **WHEN** `stream_rows_range(1000, 2000)` is called
- **THEN** only rows 1000-2000 are parsed
- **AND** rows outside the range are skipped efficiently

---

### Requirement: Lazy Workbook Mode

The system SHALL provide a lazy workbook mode that defers sheet loading.

#### Scenario: Open lazy workbook
- **GIVEN** a 300MB xlsx file
- **WHEN** `Workbook::open_lazy()` is called
- **THEN** only workbook metadata is parsed (sheet names, properties)
- **AND** sheet content is NOT loaded
- **AND** the operation completes in less than 1 second

#### Scenario: Load sheet on demand
- **GIVEN** a lazy workbook with 10 sheets
- **WHEN** `load_sheet("Sheet1")` is called
- **THEN** only Sheet1 content is parsed
- **AND** other sheets remain unloaded

---

### Requirement: Backward Compatibility

Performance optimizations SHALL NOT break existing API contracts.

#### Scenario: Existing API unchanged
- **GIVEN** code using `Workbook::open()` and `get_cell()`
- **WHEN** upgrading to the optimized version
- **THEN** the code compiles without modification
- **AND** behavior is identical (same results)

#### Scenario: Default behavior preserved
- **GIVEN** no explicit lazy/streaming API usage
- **WHEN** using standard `Workbook` API
- **THEN** all data is available after `open()` completes
- **AND** random access to any cell is immediate

---

