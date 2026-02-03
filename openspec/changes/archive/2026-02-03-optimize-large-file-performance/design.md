# Design: Large File Performance Optimization

## Context

XLEX needs to handle large xlsx files (up to 300MB+), but the current implementation uses eager loading mode, parsing all content when opening files. This violates the core architecture principles defined in `project.md`:

> - **Streaming-first**: Never materialize full tables in memory
> - **Lazy evaluation**: SharedStrings and styles loaded on-demand

## Current Architecture Analysis

### SharedStrings Parser (Problem Area)
```rust
// Current: Load all strings at once
pub fn parse_all<R: Read + BufRead>(&mut self, reader: R) -> XlexResult<Vec<String>> {
    // ... iterate through entire XML, build complete strings Vec
    self.strings = Some(strings.clone());  // Store everything in memory
    Ok(strings)
}
```

**Problem**: For files containing 1 million unique strings, this consumes significant memory and time.

### Workbook Parser (Problem Area)
```rust
// Current: Parse all sheets when opening
for (index, info) in sheet_infos.into_iter().enumerate() {
    let sheet = self.parse_sheet(/* ... */)?;  // Synchronously parse each sheet
    sheets.push(sheet);
}
```

**Problem**: Even when only reading a single cell, must wait for all sheets to be parsed.

## Proposed Architecture

### Decision 1: Lazy SharedStrings with Index

**Choice**: Build byte offset index, read individual strings on-demand.

**Design**:
```rust
pub struct LazySharedStrings {
    /// Memory-mapped file or raw bytes for the sharedStrings.xml
    data: SharedStringsData,
    /// Byte offset index: string_index -> (start_offset, length)
    index: Vec<(u64, u32)>,
    /// LRU cache for recently accessed strings
    cache: LruCache<u32, String>,
    /// Total count
    count: u32,
}

enum SharedStringsData {
    Mmap(Mmap),           // For file-based access
    InMemory(Vec<u8>),    // For stdin or small files
}
```

**Index Building Strategy**:
1. Single-pass XML scan, only record byte offset of each `<si>`
2. Don't parse actual content, just mark positions
3. When actually reading, seek to specified position and parse single `<si>`

**Rationale**:
- Index building time << full parsing time
- Memory usage: `8 bytes * string_count` (index) + LRU cache
- Supports random access reads

### Decision 2: Memory-mapped ZIP Access

**Choice**: Use `memmap2` for large files, keep original approach for small files or stdin.

**Design**:
```rust
pub struct WorkbookReader {
    source: ReaderSource,
}

enum ReaderSource {
    /// Memory-mapped file for large files (> 10MB)
    Mmap { mmap: Mmap, path: PathBuf },
    /// In-memory buffer for small files or stdin
    Buffer { data: Vec<u8>, path: Option<PathBuf> },
}

impl WorkbookReader {
    pub fn open(path: &Path) -> XlexResult<Self> {
        let metadata = std::fs::metadata(path)?;
        
        if metadata.len() > MMAP_THRESHOLD {
            // Use memory mapping for large files
            let file = File::open(path)?;
            let mmap = unsafe { Mmap::map(&file)? };
            Ok(Self { source: ReaderSource::Mmap { mmap, path: path.to_path_buf() } })
        } else {
            // Read into memory for small files
            let data = std::fs::read(path)?;
            Ok(Self { source: ReaderSource::Buffer { data, path: Some(path.to_path_buf()) } })
        }
    }
}
```

**Threshold**: 10MB (configurable)

**Rationale**:
- mmap lets OS manage page cache, avoiding double buffering
- Particularly effective for random access (reading specific sheets, specific strings)
- Small files don't need mmap overhead

### Decision 3: Parallel Sheet Parsing

**Choice**: Use `rayon` for parallel parsing of multiple sheets.

**Design**:
```rust
use rayon::prelude::*;

impl WorkbookParser {
    pub fn parse_sheets_parallel(
        &self,
        archive: &ZipArchive<impl Read + Seek>,
        sheet_infos: &[SheetInfo],
        shared_strings: &LazySharedStrings,
    ) -> XlexResult<Vec<Sheet>> {
        sheet_infos
            .par_iter()
            .map(|info| {
                // Each thread gets its own ZipArchive reader
                let sheet_data = self.read_sheet_data(archive, info)?;
                self.parse_sheet_bytes(&sheet_data, info.clone(), shared_strings)
            })
            .collect()
    }
}
```

**Considerations**:
- `ZipArchive` is not `Sync`, need to clone for each thread or use channels
- For single large sheets (millions of rows), different strategy needed (row-based parallelism)
- Keep single-threaded path as fallback

**Rationale**:
- Multi-sheet scenarios can achieve near N-times speedup
- rayon's work-stealing mechanism automatically balances load

### Decision 4: Streaming API Design

**Choice**: Add iterator-based API without affecting existing API.

**New API**:
```rust
impl Workbook {
    /// Opens a workbook in lazy mode (minimal upfront parsing)
    pub fn open_lazy(path: impl AsRef<Path>) -> XlexResult<LazyWorkbook> { ... }
    
    /// Stream rows from a sheet without loading all into memory
    pub fn stream_rows(&self, sheet: &str) -> XlexResult<RowIterator<'_>> { ... }
    
    /// Stream rows within a range
    pub fn stream_rows_range(
        &self, 
        sheet: &str, 
        start_row: u32, 
        end_row: u32
    ) -> XlexResult<RowIterator<'_>> { ... }
}

/// Lazy workbook that only parses metadata on open
pub struct LazyWorkbook {
    reader: WorkbookReader,
    metadata: WorkbookMetadata,
    // Sheets are NOT pre-loaded
}

impl LazyWorkbook {
    pub fn sheet_names(&self) -> &[String] { ... }
    pub fn load_sheet(&mut self, name: &str) -> XlexResult<&Sheet> { ... }
    pub fn stream_rows(&self, sheet: &str) -> XlexResult<RowIterator<'_>> { ... }
}

/// Iterator over rows in a sheet
pub struct RowIterator<'a> {
    xml_reader: Reader<BufReader<...>>,
    shared_strings: &'a LazySharedStrings,
    current_row: u32,
    end_row: Option<u32>,
}

impl Iterator for RowIterator<'_> {
    type Item = XlexResult<Row>;
    // ...
}
```

**Rationale**:
- New `LazyWorkbook` type clearly expresses different usage pattern
- Iterator pattern follows Rust conventions, supports `for` loops and iterator combinators
- Existing `Workbook` is unaffected

## Trade-offs

### Lazy vs Eager Loading

| Aspect | Eager (Current) | Lazy (Proposed) |
|--------|-----------------|-----------------|
| First cell access | Fast (already loaded) | Slower (on-demand parse) |
| Metadata queries | Fast | Fast (only parse workbook.xml) |
| Memory usage | O(file_size) | O(active_data) |
| Random access pattern | Fast | Fast with cache |
| Sequential full read | Same | Slightly slower (index overhead) |

**Conclusion**: For large files and partial read scenarios, lazy loading is clearly superior to eager loading.

### Memory-mapped vs Buffered I/O

| Aspect | Mmap | Buffered |
|--------|------|----------|
| Large file | Excellent | Poor (high memory) |
| Small file | Overhead | Efficient |
| stdin | Not possible | Required |
| Random access | Excellent | Good with seeking |
| Cross-platform | Minor differences | Consistent |

**Conclusion**: Use threshold-based automatic selection.

## Performance Expectations

Based on experience from similar projects:

| Operation | Current (estimated) | After Optimization |
|-----------|--------------------|--------------------|
| Open 300MB file | 30-60 seconds | < 5 seconds (index only) |
| `xlex info` | 30-60 seconds | < 1 second |
| Read single cell | < 1ms | < 10ms (with cache hit) |
| Stream all rows | N/A (OOM risk) | Streaming, constant memory |

## Migration Path

1. **Phase 1**: Lazy SharedStrings - No breaking changes
2. **Phase 2**: Memory-mapped ZIP - No breaking changes
3. **Phase 3**: Parallel parsing - No breaking changes
4. **Phase 4**: New streaming API - New functionality, doesn't affect existing

All phases can be implemented and deployed independently.

## Dependencies

New crates to add:
- `memmap2 = "0.9"` - Memory-mapped files
- `rayon = "1.10"` - Data parallelism

Both are mature, core crates in the Rust ecosystem.
