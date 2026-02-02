//! # xlex-core
//!
//! A streaming Excel manipulation library for Rust.
//!
//! `xlex-core` provides low-level access to xlsx files with a focus on
//! streaming and memory efficiency. It can handle files up to 200MB
//! without memory exhaustion.
//!
//! ## Features
//!
//! - **Streaming ZIP/XML parsing**: Process xlsx files without loading entirely into memory
//! - **Lazy SharedStrings**: On-demand string loading with LRU caching
//! - **Copy-on-write modifications**: Efficient file updates
//! - **Comprehensive cell support**: All Excel cell types and formulas
//!
//! ## Example
//!
//! ```rust,no_run
//! use xlex_core::{Workbook, CellRef};
//!
//! // Open a workbook
//! let workbook = Workbook::open("report.xlsx").unwrap();
//!
//! // Get sheet names
//! for sheet in workbook.sheet_names() {
//!     println!("Sheet: {}", sheet);
//! }
//!
//! // Read a cell
//! let cell = workbook.get_cell("Sheet1", &CellRef::new(1, 1)).unwrap();
//! println!("A1 = {:?}", cell);
//! ```

pub mod cell;
pub mod error;
pub mod parser;
pub mod range;
pub mod sheet;
pub mod style;
pub mod workbook;
pub mod writer;

// Re-exports
pub use cell::{Cell, CellRef, CellValue};
pub use error::{XlexError, XlexResult};
pub use range::Range;
pub use sheet::Sheet;
pub use style::{Style, StyleRegistry};
pub use workbook::{DefinedName, Workbook};

/// Library version
pub const VERSION: &str = env!("CARGO_PKG_VERSION");
