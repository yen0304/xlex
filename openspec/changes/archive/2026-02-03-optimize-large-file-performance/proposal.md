# Change: Optimize Large File Performance

## Why

When processing large xlsx files (300MB+), the current implementation loads all SharedStrings and all sheet cells into memory at once, resulting in:
1. Long file opening times (waiting for complete parsing)
2. High memory consumption
3. Even simple operations like `xlex info` require full parsing
4. Each CLI command re-opens the file, making multi-command workflows extremely slow

According to the project constraint "Must handle files up to 200MB without memory exhaustion", we need to optimize large file handling performance.

## What Changes

### Phase 1: Lazy SharedStrings (Internal Optimization) ✅ DONE
- Change SharedStrings to build an index and load on-demand instead of `parse_all`
- Build byte offset index using memory-mapped file for fast positioning
- Maintain LRU cache mechanism

### Phase 2: Memory-mapped ZIP Access (Internal Optimization) ✅ DONE
- Use `memmap2` crate for memory mapping large files
- Avoid repeated I/O operations
- Automatic fallback to normal read mode (for stdin or small files)

### Phase 3: Parallel Sheet Parsing (Internal Optimization) ✅ DONE
- Use `rayon` for parallel parsing of multiple sheets
- Feature flag `parallel` (default enabled)

### Phase 4: Streaming API (New API, Backward Compatible) ✅ DONE
- Add `LazyWorkbook` for metadata-only opening (sub-second for any file size)
- Add `stream_rows()` for streaming row access
- Add `read_cell()` for single cell access without loading entire sheet

### Phase 5: Session Mode (New CLI Feature)
- Add `xlex session <file>` command for interactive session with pre-loaded workbook
- Load file once, run multiple commands instantly
- Support both standard mode (full load) and lazy mode (metadata only)
- Commands: `info`, `sheet list`, `cell get`, `row get`, `exit`

## Impact

- Affected specs: None (will create new performance-related spec)
- Affected code:
  - `crates/xlex-core/src/parser/lazy_shared_strings.rs` - NEW: Lazy loading ✅
  - `crates/xlex-core/src/reader.rs` - NEW: Mmap support ✅
  - `crates/xlex-core/src/lazy.rs` - NEW: LazyWorkbook ✅
  - `crates/xlex-core/src/parser/workbook.rs` - Parallel parsing ✅
  - `crates/xlex-cli/src/commands/mod.rs` - Add session command

## Backward Compatibility

**100% Backward Compatible**:
- All existing public API signatures remain unchanged
- Existing behavior is unaffected
- New APIs are opt-in supplementary features
- Performance improvements are purely internal optimizations

## Success Criteria

| Metric | Before | After | Status |
|--------|--------|-------|--------|
| 366MB file `Workbook::open()` | 86s | < 10s | ❌ Not optimized (full parsing required) |
| 366MB file `LazyWorkbook::open()` | N/A | 0.23s | ✅ Done |
| 366MB file session mode | N/A | Load 0.23s, instant commands | ✅ Done |
| Memory usage for 300MB file | TBD | < 600MB peak | TBD |
| Existing tests | - | All pass (290 tests) | ✅ Done |
| **Speedup vs traditional** | - | **~190x faster** | ✅ Done |
