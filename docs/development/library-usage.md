# Library Usage

xlex-core can be used as a Rust library for programmatic Excel manipulation.

## Installation

Add to your `Cargo.toml`:

```toml
[dependencies]
xlex-core = "0.3"
```

## Quick Start

```rust
use xlex_core::{Workbook, CellRef, CellValue};

// Open a workbook
let workbook = Workbook::open("report.xlsx")?;

// List sheets
for name in workbook.sheet_names() {
    println!("Sheet: {}", name);
}

// Read a cell
let value = workbook.get_cell("Sheet1", &CellRef::new(1, 1))?;
println!("A1 = {}", value.to_display_string());
```

## Creating a Workbook

```rust
use xlex_core::{Workbook, CellRef, CellValue};

let mut workbook = Workbook::new();

// Add data
if let Some(sheet) = workbook.get_sheet_mut("Sheet1") {
    sheet.set_cell(CellRef::new(1, 1), CellValue::string("Hello"));
    sheet.set_cell(CellRef::new(2, 1), CellValue::number(42.0));
}

workbook.save_as("output.xlsx")?;
```

## Lazy Loading for Large Files

For files over 10MB, use `LazyWorkbook` for memory-efficient access:

```rust
use xlex_core::LazyWorkbook;

let lazy = LazyWorkbook::open("large_file.xlsx")?;

// Stream rows without loading entire file
for row in lazy.stream_rows("Sheet1")? {
    println!("{:?}", row);
}
```

## Key Types

| Type | Description |
|------|-------------|
| `Workbook` | Full in-memory workbook |
| `LazyWorkbook` | Streaming/lazy workbook for large files |
| `CellRef` | Cell reference (column, row) |
| `CellValue` | Cell value (String, Number, Boolean, Formula, etc.) |
| `Style` | Cell formatting (font, fill, border, alignment) |

For the complete API, see the [docs.rs documentation](https://docs.rs/xlex-core).
