//! XLSX parsing utilities.

mod shared_strings;
mod styles;
mod workbook;

pub use shared_strings::SharedStringsParser;
pub use styles::StylesParser;
pub use workbook::WorkbookParser;

use crate::error::{XlexError, XlexResult};

/// Required entries in an xlsx file.
pub const REQUIRED_ENTRIES: &[&str] = &[
    "[Content_Types].xml",
    "xl/workbook.xml",
];

/// Validates that a ZIP archive contains required xlsx entries.
pub fn validate_xlsx_structure<R: std::io::Read + std::io::Seek>(
    archive: &mut zip::ZipArchive<R>,
) -> XlexResult<()> {
    for entry in REQUIRED_ENTRIES {
        if archive.by_name(entry).is_err() {
            return Err(XlexError::MissingRequiredEntry {
                entry: entry.to_string(),
            });
        }
    }
    Ok(())
}
