//! XLSX parsing utilities.

mod lazy_shared_strings;
mod shared_strings;
mod styles;
mod workbook;

pub use lazy_shared_strings::LazySharedStrings;
pub use shared_strings::SharedStringsParser;
pub use styles::StylesParser;
pub use workbook::WorkbookParser;

use crate::error::{XlexError, XlexResult};

/// Required entries in an xlsx file.
pub const REQUIRED_ENTRIES: &[&str] = &["[Content_Types].xml", "xl/workbook.xml"];

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

#[cfg(test)]
mod tests {
    use super::*;
    use std::io::Cursor;

    #[test]
    fn test_required_entries() {
        assert!(REQUIRED_ENTRIES.contains(&"[Content_Types].xml"));
        assert!(REQUIRED_ENTRIES.contains(&"xl/workbook.xml"));
    }

    #[test]
    fn test_validate_xlsx_structure_valid() {
        // Create a minimal valid xlsx structure
        let mut buf = Vec::new();
        {
            let mut zip = zip::ZipWriter::new(Cursor::new(&mut buf));
            let options = zip::write::SimpleFileOptions::default();

            zip.start_file("[Content_Types].xml", options).unwrap();
            zip.write_all(b"<?xml version=\"1.0\"?><Types/>").unwrap();

            zip.start_file("xl/workbook.xml", options).unwrap();
            zip.write_all(b"<?xml version=\"1.0\"?><workbook/>")
                .unwrap();

            zip.finish().unwrap();
        }

        let cursor = Cursor::new(buf);
        let mut archive = zip::ZipArchive::new(cursor).unwrap();

        let result = validate_xlsx_structure(&mut archive);
        assert!(result.is_ok());
    }

    #[test]
    fn test_validate_xlsx_structure_missing_content_types() {
        // Create zip without [Content_Types].xml
        let mut buf = Vec::new();
        {
            let mut zip = zip::ZipWriter::new(Cursor::new(&mut buf));
            let options = zip::write::SimpleFileOptions::default();

            zip.start_file("xl/workbook.xml", options).unwrap();
            zip.write_all(b"<?xml version=\"1.0\"?><workbook/>")
                .unwrap();

            zip.finish().unwrap();
        }

        let cursor = Cursor::new(buf);
        let mut archive = zip::ZipArchive::new(cursor).unwrap();

        let result = validate_xlsx_structure(&mut archive);
        assert!(result.is_err());
        assert!(matches!(
            result.unwrap_err(),
            XlexError::MissingRequiredEntry { .. }
        ));
    }

    #[test]
    fn test_validate_xlsx_structure_missing_workbook() {
        // Create zip without xl/workbook.xml
        let mut buf = Vec::new();
        {
            let mut zip = zip::ZipWriter::new(Cursor::new(&mut buf));
            let options = zip::write::SimpleFileOptions::default();

            zip.start_file("[Content_Types].xml", options).unwrap();
            zip.write_all(b"<?xml version=\"1.0\"?><Types/>").unwrap();

            zip.finish().unwrap();
        }

        let cursor = Cursor::new(buf);
        let mut archive = zip::ZipArchive::new(cursor).unwrap();

        let result = validate_xlsx_structure(&mut archive);
        assert!(result.is_err());
        assert!(matches!(
            result.unwrap_err(),
            XlexError::MissingRequiredEntry { .. }
        ));
    }

    use std::io::Write;
}
