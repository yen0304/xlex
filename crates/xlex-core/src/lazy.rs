//! Lazy workbook for streaming large files.
//!
//! This module provides a lazy workbook implementation that defers parsing
//! until data is actually needed, enabling efficient handling of large files.

use std::collections::HashMap;
use std::io::{BufReader, Cursor, Read, Seek};
use std::path::Path;
use std::sync::{Arc, Mutex};

use quick_xml::events::Event;
use quick_xml::Reader;
use zip::ZipArchive;

use crate::cell::{CellError, CellRef, CellValue};
use crate::error::{XlexError, XlexResult};
use crate::parser::{validate_xlsx_structure, LazySharedStrings};
use crate::reader::WorkbookReader;
use crate::sheet::SheetInfo;

/// A lazy workbook that only parses sheets on demand.
///
/// Unlike [`Workbook`](crate::Workbook), this struct does not load all sheets
/// into memory at once. Instead, it provides streaming access to sheet data.
///
/// # Example
///
/// ```no_run
/// use xlex_core::LazyWorkbook;
///
/// let wb = LazyWorkbook::open("large_file.xlsx")?;
/// println!("Sheets: {:?}", wb.sheet_names());
///
/// // Stream rows from a specific sheet
/// for row in wb.stream_rows("Sheet1")? {
///     println!("Row {}: {:?}", row.row_number, row.cells);
/// }
/// # Ok::<(), xlex_core::XlexError>(())
/// ```
pub struct LazyWorkbook {
    /// Raw workbook data
    data: Arc<Vec<u8>>,
    /// Sheet metadata (name -> (index, info, zip_path))
    sheets: HashMap<String, (usize, SheetInfo, String)>,
    /// Ordered sheet names
    sheet_names: Vec<String>,
    /// Shared strings (lazy loaded, wrapped in Mutex for interior mutability)
    shared_strings: Arc<Mutex<LazySharedStrings>>,
}

/// A row from a sheet being streamed.
#[derive(Debug, Clone)]
pub struct StreamRow {
    /// 1-based row number
    pub row_number: u32,
    /// Cells in this row (may be sparse)
    pub cells: Vec<(CellRef, CellValue)>,
}

impl LazyWorkbook {
    /// Opens a workbook lazily from a file path.
    ///
    /// This is very fast as it only reads metadata, not sheet contents.
    pub fn open(path: impl AsRef<Path>) -> XlexResult<Self> {
        let path = path.as_ref();

        // Check extension
        if path.extension().and_then(|e| e.to_str()) != Some("xlsx") {
            return Err(XlexError::InvalidExtension {
                path: path.to_path_buf(),
            });
        }

        let reader = WorkbookReader::open(path)?;
        Self::from_bytes(reader.as_bytes().to_vec())
    }

    /// Creates a lazy workbook from raw bytes.
    pub fn from_bytes(data: Vec<u8>) -> XlexResult<Self> {
        let data = Arc::new(data);

        // Quick validation
        {
            let cursor = Cursor::new(data.as_ref());
            let mut archive = ZipArchive::new(cursor)?;
            validate_xlsx_structure(&mut archive)?;
        }

        // Parse shared strings lazily
        let shared_strings = {
            let cursor = Cursor::new(data.as_ref());
            let mut archive = ZipArchive::new(cursor)?;

            let result = if let Ok(mut file) = archive.by_name("xl/sharedStrings.xml") {
                let mut ss_data = Vec::new();
                file.read_to_end(&mut ss_data)
                    .map_err(|e| XlexError::IoError {
                        message: e.to_string(),
                        source: Some(e),
                    })?;
                Arc::new(Mutex::new(LazySharedStrings::from_bytes_default(ss_data)?))
            } else {
                Arc::new(Mutex::new(LazySharedStrings::default()))
            };
            result
        };

        // Parse sheet metadata only
        let (sheets, sheet_names) = {
            let cursor = Cursor::new(data.as_ref());
            let mut archive = ZipArchive::new(cursor)?;
            Self::parse_sheet_metadata(&mut archive)?
        };

        Ok(Self {
            data,
            sheets,
            sheet_names,
            shared_strings,
        })
    }

    /// Returns the list of sheet names.
    pub fn sheet_names(&self) -> &[String] {
        &self.sheet_names
    }

    /// Returns the number of sheets.
    pub fn sheet_count(&self) -> usize {
        self.sheet_names.len()
    }

    /// Checks if a sheet exists.
    pub fn has_sheet(&self, name: &str) -> bool {
        self.sheets.contains_key(name)
    }

    /// Streams all rows from a sheet and returns them as a vector.
    ///
    /// This is the primary method for reading sheet data in streaming fashion.
    /// Each row is parsed on demand from the underlying XML.
    pub fn stream_rows(&self, sheet_name: &str) -> XlexResult<Vec<StreamRow>> {
        let (_index, _info, zip_path) =
            self.sheets
                .get(sheet_name)
                .ok_or_else(|| XlexError::SheetNotFound {
                    name: sheet_name.to_string(),
                })?;

        let cursor = Cursor::new(self.data.as_ref().as_slice());
        let mut archive = ZipArchive::new(cursor)?;
        let file = archive.by_name(zip_path)?;

        self.parse_rows_from_sheet(BufReader::new(file))
    }

    /// Reads a single cell value without loading the entire sheet.
    pub fn read_cell(&self, sheet_name: &str, cell_ref: &CellRef) -> XlexResult<Option<CellValue>> {
        let (_index, _info, zip_path) =
            self.sheets
                .get(sheet_name)
                .ok_or_else(|| XlexError::SheetNotFound {
                    name: sheet_name.to_string(),
                })?;

        let cursor = Cursor::new(self.data.as_ref().as_slice());
        let mut archive = ZipArchive::new(cursor)?;
        let file = archive.by_name(zip_path)?;

        self.find_cell_in_sheet(BufReader::new(file), cell_ref)
    }

    /// Parses only the sheet metadata (names, paths) without loading content.
    #[allow(clippy::type_complexity)]
    fn parse_sheet_metadata<R: Read + Seek>(
        archive: &mut ZipArchive<R>,
    ) -> XlexResult<(HashMap<String, (usize, SheetInfo, String)>, Vec<String>)> {
        // Parse relationships to get sheet file paths
        let relationships = Self::parse_relationships(archive)?;

        // Parse workbook.xml for sheet info
        let file = archive.by_name("xl/workbook.xml")?;
        let mut reader = Reader::from_reader(BufReader::new(file));
        reader.config_mut().trim_text(true);

        let mut sheets = HashMap::new();
        let mut sheet_names = Vec::new();
        let mut buf = Vec::new();
        let mut index = 0;

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Empty(e)) | Ok(Event::Start(e)) if e.name().as_ref() == b"sheet" => {
                    let mut name = String::new();
                    let mut rel_id = String::new();
                    let mut sheet_id = 0u32;

                    for attr in e.attributes().flatten() {
                        match attr.key.as_ref() {
                            b"name" => {
                                name = String::from_utf8_lossy(&attr.value).to_string();
                            }
                            b"sheetId" => {
                                sheet_id =
                                    String::from_utf8_lossy(&attr.value).parse().unwrap_or(0);
                            }
                            b"r:id" => {
                                rel_id = String::from_utf8_lossy(&attr.value).to_string();
                            }
                            _ => {}
                        }
                    }

                    let zip_path = relationships
                        .get(&rel_id)
                        .map(|s| format!("xl/{}", s))
                        .unwrap_or_else(|| format!("xl/worksheets/sheet{}.xml", index + 1));

                    let info = SheetInfo::new(name.clone(), sheet_id, rel_id, index);

                    sheet_names.push(name.clone());
                    sheets.insert(name, (index, info, zip_path));
                    index += 1;
                }
                Ok(Event::Eof) => break,
                Err(e) => {
                    return Err(XlexError::InvalidXml {
                        message: e.to_string(),
                    })
                }
                _ => {}
            }
            buf.clear();
        }

        Ok((sheets, sheet_names))
    }

    /// Parses relationships from workbook.xml.rels.
    fn parse_relationships<R: Read + Seek>(
        archive: &mut ZipArchive<R>,
    ) -> XlexResult<HashMap<String, String>> {
        let mut relationships = HashMap::new();

        if let Ok(file) = archive.by_name("xl/_rels/workbook.xml.rels") {
            let mut reader = Reader::from_reader(BufReader::new(file));
            reader.config_mut().trim_text(true);

            let mut buf = Vec::new();
            loop {
                match reader.read_event_into(&mut buf) {
                    Ok(Event::Empty(e)) if e.name().as_ref() == b"Relationship" => {
                        let mut id = String::new();
                        let mut target = String::new();

                        for attr in e.attributes().flatten() {
                            match attr.key.as_ref() {
                                b"Id" => {
                                    id = String::from_utf8_lossy(&attr.value).to_string();
                                }
                                b"Target" => {
                                    target = String::from_utf8_lossy(&attr.value).to_string();
                                }
                                _ => {}
                            }
                        }

                        if !id.is_empty() && !target.is_empty() {
                            relationships.insert(id, target);
                        }
                    }
                    Ok(Event::Eof) => break,
                    Err(_) => break,
                    _ => {}
                }
                buf.clear();
            }
        }

        Ok(relationships)
    }

    /// Finds a specific cell in a sheet without loading the entire sheet.
    fn find_cell_in_sheet<R: Read>(
        &self,
        reader: R,
        target: &CellRef,
    ) -> XlexResult<Option<CellValue>> {
        let target_ref = target.to_a1();
        let mut xml_reader = Reader::from_reader(BufReader::new(reader));
        xml_reader.config_mut().trim_text(true);

        let mut buf = Vec::new();
        let mut in_target_cell = false;
        let mut cell_type = String::new();
        let mut in_value = false;
        let mut value_text = String::new();

        loop {
            match xml_reader.read_event_into(&mut buf) {
                Ok(Event::Start(e)) => match e.name().as_ref() {
                    b"c" => {
                        // Check if this is our target cell
                        for attr in e.attributes().flatten() {
                            if attr.key.as_ref() == b"r" {
                                let r = String::from_utf8_lossy(&attr.value);
                                if r == target_ref {
                                    in_target_cell = true;
                                    // Get cell type
                                    for attr2 in e.attributes().flatten() {
                                        if attr2.key.as_ref() == b"t" {
                                            cell_type =
                                                String::from_utf8_lossy(&attr2.value).to_string();
                                        }
                                    }
                                }
                            }
                        }
                    }
                    b"v" if in_target_cell => {
                        in_value = true;
                    }
                    _ => {}
                },
                Ok(Event::Text(e)) if in_value => {
                    value_text = e.unescape().unwrap_or_default().to_string();
                }
                Ok(Event::End(e)) => match e.name().as_ref() {
                    b"c" if in_target_cell => {
                        // We found our cell, convert and return
                        let value = self.convert_cell_value(&cell_type, &value_text)?;
                        return Ok(Some(value));
                    }
                    b"v" => {
                        in_value = false;
                    }
                    b"row" if in_target_cell => {
                        // Passed our target row without finding, cell is empty
                        return Ok(None);
                    }
                    _ => {}
                },
                Ok(Event::Eof) => break,
                Err(e) => {
                    return Err(XlexError::InvalidXml {
                        message: e.to_string(),
                    })
                }
                _ => {}
            }
            buf.clear();
        }

        Ok(None)
    }

    /// Converts raw cell value based on type.
    fn convert_cell_value(&self, cell_type: &str, value: &str) -> XlexResult<CellValue> {
        if value.is_empty() {
            return Ok(CellValue::Empty);
        }

        match cell_type {
            "s" => {
                // Shared string index
                let idx: u32 = value.parse().map_err(|_| XlexError::InvalidXml {
                    message: format!("Invalid shared string index: {}", value),
                })?;
                let mut ss = self.shared_strings.lock().unwrap();
                if let Some(s) = ss.get(idx) {
                    Ok(CellValue::String(s))
                } else {
                    Ok(CellValue::String(String::new()))
                }
            }
            "b" => {
                // Boolean
                Ok(CellValue::Boolean(
                    value == "1" || value.eq_ignore_ascii_case("true"),
                ))
            }
            "e" => {
                // Error - use CellError::parse or default to Value
                let error = CellError::parse(value).unwrap_or(CellError::Value);
                Ok(CellValue::Error(error))
            }
            "str" | "inlineStr" => {
                // Inline string
                Ok(CellValue::String(value.to_string()))
            }
            _ => {
                // Number (default)
                if let Ok(n) = value.parse::<f64>() {
                    Ok(CellValue::Number(n))
                } else {
                    Ok(CellValue::String(value.to_string()))
                }
            }
        }
    }

    /// Parses all rows from a sheet reader.
    fn parse_rows_from_sheet<R: Read>(&self, reader: R) -> XlexResult<Vec<StreamRow>> {
        let mut xml_reader = Reader::from_reader(BufReader::new(reader));
        xml_reader.config_mut().trim_text(true);

        let mut rows = Vec::new();
        let mut buf = Vec::new();
        let mut current_row: Option<u32> = None;
        let mut current_cells: Vec<(CellRef, CellValue)> = Vec::new();

        // Cell parsing state
        let mut cell_ref: Option<CellRef> = None;
        let mut cell_type = String::new();
        let mut in_value = false;
        let mut value_text = String::new();

        loop {
            match xml_reader.read_event_into(&mut buf) {
                Ok(Event::Start(e)) => match e.name().as_ref() {
                    b"row" => {
                        // Get row number
                        for attr in e.attributes().flatten() {
                            if attr.key.as_ref() == b"r" {
                                current_row = String::from_utf8_lossy(&attr.value).parse().ok();
                            }
                        }
                        current_cells.clear();
                    }
                    b"c" => {
                        cell_type.clear();
                        cell_ref = None;

                        for attr in e.attributes().flatten() {
                            match attr.key.as_ref() {
                                b"r" => {
                                    let r = String::from_utf8_lossy(&attr.value);
                                    cell_ref = r.parse().ok();
                                }
                                b"t" => {
                                    cell_type = String::from_utf8_lossy(&attr.value).to_string();
                                }
                                _ => {}
                            }
                        }
                    }
                    b"v" => {
                        in_value = true;
                        value_text.clear();
                    }
                    _ => {}
                },
                Ok(Event::Text(e)) if in_value => {
                    value_text = e.unescape().unwrap_or_default().to_string();
                }
                Ok(Event::End(e)) => match e.name().as_ref() {
                    b"v" => {
                        in_value = false;
                    }
                    b"c" => {
                        if let Some(ref_) = cell_ref.take() {
                            let value = self
                                .convert_cell_value(&cell_type, &value_text)
                                .unwrap_or(CellValue::Empty);
                            current_cells.push((ref_, value));
                        }
                        value_text.clear();
                    }
                    b"row" => {
                        if let Some(row_num) = current_row.take() {
                            rows.push(StreamRow {
                                row_number: row_num,
                                cells: std::mem::take(&mut current_cells),
                            });
                        }
                    }
                    _ => {}
                },
                Ok(Event::Empty(e)) if e.name().as_ref() == b"c" => {
                    // Empty cell element
                    for attr in e.attributes().flatten() {
                        if attr.key.as_ref() == b"r" {
                            let r = String::from_utf8_lossy(&attr.value);
                            if let Ok(ref_) = r.parse::<CellRef>() {
                                current_cells.push((ref_, CellValue::Empty));
                            }
                        }
                    }
                }
                Ok(Event::Eof) => break,
                Err(e) => {
                    return Err(XlexError::InvalidXml {
                        message: e.to_string(),
                    })
                }
                _ => {}
            }
            buf.clear();
        }

        Ok(rows)
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_lazy_workbook_from_bytes_empty() {
        // Test with invalid data
        let result = LazyWorkbook::from_bytes(vec![0, 1, 2, 3]);
        assert!(result.is_err());
    }

    #[test]
    fn test_stream_row_debug() {
        let row = StreamRow {
            row_number: 1,
            cells: vec![],
        };
        assert!(format!("{:?}", row).contains("row_number: 1"));
    }

    #[test]
    fn test_convert_cell_value_empty() {
        let wb = LazyWorkbook {
            data: Arc::new(vec![]),
            sheets: HashMap::new(),
            sheet_names: vec![],
            shared_strings: Arc::new(Mutex::new(LazySharedStrings::default())),
        };

        let value = wb.convert_cell_value("", "").unwrap();
        assert!(matches!(value, CellValue::Empty));
    }

    #[test]
    fn test_convert_cell_value_boolean() {
        let wb = LazyWorkbook {
            data: Arc::new(vec![]),
            sheets: HashMap::new(),
            sheet_names: vec![],
            shared_strings: Arc::new(Mutex::new(LazySharedStrings::default())),
        };

        let value = wb.convert_cell_value("b", "1").unwrap();
        assert!(matches!(value, CellValue::Boolean(true)));

        let value = wb.convert_cell_value("b", "0").unwrap();
        assert!(matches!(value, CellValue::Boolean(false)));
    }

    #[test]
    fn test_convert_cell_value_number() {
        let wb = LazyWorkbook {
            data: Arc::new(vec![]),
            sheets: HashMap::new(),
            sheet_names: vec![],
            shared_strings: Arc::new(Mutex::new(LazySharedStrings::default())),
        };

        let value = wb.convert_cell_value("n", "42.5").unwrap();
        assert!(matches!(value, CellValue::Number(n) if (n - 42.5).abs() < f64::EPSILON));
    }

    #[test]
    fn test_convert_cell_value_error() {
        let wb = LazyWorkbook {
            data: Arc::new(vec![]),
            sheets: HashMap::new(),
            sheet_names: vec![],
            shared_strings: Arc::new(Mutex::new(LazySharedStrings::default())),
        };

        let value = wb.convert_cell_value("e", "#DIV/0!").unwrap();
        assert!(matches!(value, CellValue::Error(CellError::DivZero)));
    }
}
