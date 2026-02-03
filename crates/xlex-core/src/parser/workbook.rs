//! Workbook parser.

use std::collections::HashMap;
use std::io::{BufRead, BufReader, Read, Seek};
use std::path::PathBuf;

use quick_xml::events::Event;
use quick_xml::Reader;
use zip::ZipArchive;

use crate::cell::{Cell, CellError, CellRef, CellValue};
use crate::error::{XlexError, XlexResult};
use crate::parser::{validate_xlsx_structure, SharedStringsParser, StylesParser};
use crate::sheet::{Sheet, SheetInfo, SheetVisibility};
use crate::style::StyleRegistry;
use crate::workbook::{DefinedName, DocumentProperties, Workbook};

/// Parser for xlsx workbooks.
pub struct WorkbookParser {
    /// Shared strings parser
    #[allow(dead_code)]
    shared_strings: SharedStringsParser,
    /// Styles parser
    styles_parser: StylesParser,
}

impl WorkbookParser {
    /// Creates a new workbook parser.
    pub fn new() -> Self {
        Self {
            shared_strings: SharedStringsParser::with_default_cache(),
            styles_parser: StylesParser::new(),
        }
    }

    /// Parses a workbook from a ZIP archive.
    pub fn parse<R: Read + Seek>(
        &self,
        archive: &mut ZipArchive<R>,
        path: Option<PathBuf>,
    ) -> XlexResult<Workbook> {
        // Validate structure
        validate_xlsx_structure(archive)?;

        // Parse shared strings if present
        let shared_strings = if let Ok(file) = archive.by_name("xl/sharedStrings.xml") {
            let mut parser = SharedStringsParser::with_default_cache();
            parser.parse_all(BufReader::new(file))?
        } else {
            Vec::new()
        };

        // Parse styles if present
        let style_registry = if let Ok(file) = archive.by_name("xl/styles.xml") {
            self.styles_parser.parse(BufReader::new(file))?
        } else {
            StyleRegistry::new()
        };

        // Parse document properties
        let properties = self.parse_properties(archive)?;

        // Parse workbook.xml to get sheet info and defined names
        let (sheet_infos, defined_names) = self.parse_workbook_xml_full(archive)?;

        // Parse relationships to get sheet file paths
        let relationships = self.parse_relationships(archive)?;

        // Parse each sheet
        let mut sheets = Vec::new();
        let mut sheet_map = HashMap::new();

        for (index, info) in sheet_infos.into_iter().enumerate() {
            let sheet_path = relationships
                .get(&info.rel_id)
                .map(|s| format!("xl/{}", s))
                .unwrap_or_else(|| format!("xl/worksheets/sheet{}.xml", index + 1));

            let sheet = if let Ok(file) = archive.by_name(&sheet_path) {
                self.parse_sheet(BufReader::new(file), info.clone(), &shared_strings)?
            } else {
                Sheet::new(info.clone())
            };

            sheet_map.insert(info.name.clone(), index);
            sheets.push(sheet);
        }

        // Construct workbook using the internal constructor
        Ok(Workbook::__from_parts(
            path,
            properties,
            sheets,
            sheet_map,
            style_registry,
            shared_strings,
            defined_names,
            0,
            false,
        ))
    }

    /// Parses document properties from core.xml and app.xml.
    fn parse_properties<R: Read + Seek>(
        &self,
        archive: &mut ZipArchive<R>,
    ) -> XlexResult<DocumentProperties> {
        let mut props = DocumentProperties::default();

        // Parse core.xml for basic properties
        if let Ok(file) = archive.by_name("docProps/core.xml") {
            let mut reader = Reader::from_reader(BufReader::new(file));
            reader.config_mut().trim_text(true);

            let mut buf = Vec::new();
            let mut current_element = String::new();

            loop {
                match reader.read_event_into(&mut buf) {
                    Ok(Event::Start(e)) => {
                        current_element = String::from_utf8_lossy(e.name().as_ref()).to_string();
                    }
                    Ok(Event::Text(e)) => {
                        let text = e.unescape().unwrap_or_default().to_string();
                        match current_element.as_str() {
                            "dc:title" | "title" => props.title = Some(text),
                            "dc:subject" | "subject" => props.subject = Some(text),
                            "dc:creator" | "creator" => props.creator = Some(text),
                            "cp:keywords" | "keywords" => props.keywords = Some(text),
                            "dc:description" | "description" => props.description = Some(text),
                            "cp:lastModifiedBy" | "lastModifiedBy" => {
                                props.last_modified_by = Some(text)
                            }
                            "cp:category" | "category" => props.category = Some(text),
                            "cp:contentStatus" | "contentStatus" => {
                                props.content_status = Some(text)
                            }
                            _ => {}
                        }
                    }
                    Ok(Event::End(_)) => {
                        current_element.clear();
                    }
                    Ok(Event::Eof) => break,
                    Err(_) => break,
                    _ => {}
                }
                buf.clear();
            }
        }

        Ok(props)
    }

    /// Parses workbook.xml to extract sheet information and defined names.
    fn parse_workbook_xml_full<R: Read + Seek>(
        &self,
        archive: &mut ZipArchive<R>,
    ) -> XlexResult<(Vec<SheetInfo>, Vec<DefinedName>)> {
        let file = archive.by_name("xl/workbook.xml")?;
        let mut reader = Reader::from_reader(BufReader::new(file));
        reader.config_mut().trim_text(true);

        let mut sheets = Vec::new();
        let mut defined_names = Vec::new();
        let mut buf = Vec::new();
        let mut in_defined_name = false;
        let mut current_defined_name: Option<DefinedName> = None;

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Empty(e)) | Ok(Event::Start(e)) if e.name().as_ref() == b"sheet" => {
                    let mut name = String::new();
                    let mut sheet_id: u32 = 0;
                    let mut rel_id = String::new();
                    let mut visibility = SheetVisibility::Visible;

                    for attr in e.attributes().flatten() {
                        match attr.key.as_ref() {
                            b"name" => {
                                name = String::from_utf8_lossy(&attr.value).to_string();
                            }
                            b"sheetId" => {
                                sheet_id =
                                    String::from_utf8_lossy(&attr.value).parse().unwrap_or(0);
                            }
                            b"state" => {
                                let state = String::from_utf8_lossy(&attr.value);
                                visibility = match state.as_ref() {
                                    "hidden" => SheetVisibility::Hidden,
                                    "veryHidden" => SheetVisibility::VeryHidden,
                                    _ => SheetVisibility::Visible,
                                };
                            }
                            key if key.ends_with(b"id") => {
                                rel_id = String::from_utf8_lossy(&attr.value).to_string();
                            }
                            _ => {}
                        }
                    }

                    if !name.is_empty() {
                        let mut info = SheetInfo::new(&name, sheet_id, &rel_id, sheets.len());
                        info.visibility = visibility;
                        sheets.push(info);
                    }
                }
                Ok(Event::Start(e)) if e.name().as_ref() == b"definedName" => {
                    in_defined_name = true;
                    let mut name = String::new();
                    let mut local_sheet_id: Option<usize> = None;
                    let mut comment: Option<String> = None;
                    let mut hidden = false;

                    for attr in e.attributes().flatten() {
                        match attr.key.as_ref() {
                            b"name" => {
                                name = String::from_utf8_lossy(&attr.value).to_string();
                            }
                            b"localSheetId" => {
                                local_sheet_id = String::from_utf8_lossy(&attr.value).parse().ok();
                            }
                            b"comment" => {
                                comment = Some(String::from_utf8_lossy(&attr.value).to_string());
                            }
                            b"hidden" => {
                                hidden = String::from_utf8_lossy(&attr.value) == "1";
                            }
                            _ => {}
                        }
                    }

                    current_defined_name = Some(DefinedName {
                        name,
                        reference: String::new(),
                        local_sheet_id,
                        comment,
                        hidden,
                    });
                }
                Ok(Event::Text(e)) if in_defined_name => {
                    if let Some(ref mut dn) = current_defined_name {
                        dn.reference = e.unescape().unwrap_or_default().to_string();
                    }
                }
                Ok(Event::End(e)) if e.name().as_ref() == b"definedName" => {
                    in_defined_name = false;
                    if let Some(dn) = current_defined_name.take() {
                        if !dn.name.is_empty() && !dn.reference.is_empty() {
                            defined_names.push(dn);
                        }
                    }
                }
                Ok(Event::Eof) => break,
                Err(e) => {
                    return Err(XlexError::InvalidXml {
                        message: format!("Error parsing workbook.xml: {}", e),
                    });
                }
                _ => {}
            }
            buf.clear();
        }

        if sheets.is_empty() {
            return Err(XlexError::ParseError {
                message: "No sheets found in workbook".to_string(),
                location: "xl/workbook.xml".to_string(),
            });
        }

        Ok((sheets, defined_names))
    }

    /// Parses relationships from xl/_rels/workbook.xml.rels.
    fn parse_relationships<R: Read + Seek>(
        &self,
        archive: &mut ZipArchive<R>,
    ) -> XlexResult<HashMap<String, String>> {
        let mut relationships = HashMap::new();

        let file = match archive.by_name("xl/_rels/workbook.xml.rels") {
            Ok(f) => f,
            Err(_) => return Ok(relationships),
        };

        let mut reader = Reader::from_reader(BufReader::new(file));
        reader.config_mut().trim_text(true);

        let mut buf = Vec::new();

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Empty(e)) | Ok(Event::Start(e))
                    if e.name().as_ref() == b"Relationship" =>
                {
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

        Ok(relationships)
    }

    /// Parses a worksheet XML file.
    fn parse_sheet<R: Read + BufRead>(
        &self,
        reader: R,
        info: SheetInfo,
        shared_strings: &[String],
    ) -> XlexResult<Sheet> {
        let mut sheet = Sheet::new(info);
        let mut xml_reader = Reader::from_reader(reader);
        xml_reader.config_mut().trim_text(true);

        let mut buf = Vec::new();
        let mut current_cell_ref: Option<CellRef> = None;
        let mut current_cell_type: Option<String> = None;
        let mut current_cell_style: Option<u32> = None;
        let mut current_value = String::new();
        let mut current_formula = String::new();
        let mut in_value = false;
        let mut in_formula = false;

        loop {
            match xml_reader.read_event_into(&mut buf) {
                Ok(Event::Start(e)) | Ok(Event::Empty(e)) => {
                    match e.name().as_ref() {
                        b"c" => {
                            // Cell element
                            for attr in e.attributes().flatten() {
                                match attr.key.as_ref() {
                                    b"r" => {
                                        let ref_str = String::from_utf8_lossy(&attr.value);
                                        current_cell_ref = CellRef::parse(&ref_str).ok();
                                    }
                                    b"t" => {
                                        current_cell_type =
                                            Some(String::from_utf8_lossy(&attr.value).to_string());
                                    }
                                    b"s" => {
                                        current_cell_style =
                                            String::from_utf8_lossy(&attr.value).parse().ok();
                                    }
                                    _ => {}
                                }
                            }
                        }
                        b"v" => {
                            in_value = true;
                            current_value.clear();
                        }
                        b"f" => {
                            in_formula = true;
                            current_formula.clear();
                        }
                        b"row" => {
                            // Could parse row attributes (height, hidden) here
                        }
                        b"col" => {
                            // Could parse column attributes (width, hidden) here
                        }
                        _ => {}
                    }
                }
                Ok(Event::Text(e)) => {
                    if in_value {
                        current_value.push_str(&e.unescape().unwrap_or_default());
                    } else if in_formula {
                        current_formula.push_str(&e.unescape().unwrap_or_default());
                    }
                }
                Ok(Event::End(e)) => {
                    match e.name().as_ref() {
                        b"c" => {
                            // Finalize cell
                            if let Some(ref cell_ref) = current_cell_ref {
                                let value = self.parse_cell_value(
                                    &current_value,
                                    &current_formula,
                                    current_cell_type.as_deref(),
                                    shared_strings,
                                );

                                let mut cell = Cell::new(cell_ref.clone(), value);
                                if let Some(style_id) = current_cell_style {
                                    cell = cell.with_style(style_id);
                                }

                                sheet.set_cell(cell_ref.clone(), cell.value);
                            }

                            // Reset state
                            current_cell_ref = None;
                            current_cell_type = None;
                            current_cell_style = None;
                            current_value.clear();
                            current_formula.clear();
                        }
                        b"v" => {
                            in_value = false;
                        }
                        b"f" => {
                            in_formula = false;
                        }
                        _ => {}
                    }
                }
                Ok(Event::Eof) => break,
                Err(e) => {
                    return Err(XlexError::InvalidXml {
                        message: format!("Error parsing sheet: {}", e),
                    });
                }
                _ => {}
            }
            buf.clear();
        }

        Ok(sheet)
    }

    /// Parses a cell value based on its type and content.
    fn parse_cell_value(
        &self,
        value: &str,
        formula: &str,
        cell_type: Option<&str>,
        shared_strings: &[String],
    ) -> CellValue {
        // If there's a formula, return formula value
        if !formula.is_empty() {
            return CellValue::Formula {
                formula: formula.to_string(),
                cached_result: if !value.is_empty() {
                    Some(Box::new(CellValue::String(value.to_string())))
                } else {
                    None
                },
            };
        }

        // Parse based on type
        match cell_type {
            Some("s") => {
                // Shared string
                if let Ok(index) = value.parse::<usize>() {
                    if let Some(s) = shared_strings.get(index) {
                        return CellValue::String(s.clone());
                    }
                }
                CellValue::Empty
            }
            Some("b") => {
                // Boolean
                CellValue::Boolean(value == "1" || value.eq_ignore_ascii_case("true"))
            }
            Some("e") => {
                // Error
                CellError::parse(value)
                    .map(CellValue::Error)
                    .unwrap_or(CellValue::Empty)
            }
            Some("str") | Some("inlineStr") => {
                // Inline string
                CellValue::String(value.to_string())
            }
            _ => {
                // Number or empty
                if value.is_empty() {
                    CellValue::Empty
                } else if let Ok(n) = value.parse::<f64>() {
                    CellValue::Number(n)
                } else {
                    CellValue::String(value.to_string())
                }
            }
        }
    }
}

impl Default for WorkbookParser {
    fn default() -> Self {
        Self::new()
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_parse_cell_value_number() {
        let parser = WorkbookParser::new();
        let value = parser.parse_cell_value("42.5", "", None, &[]);
        assert_eq!(value, CellValue::Number(42.5));
    }

    #[test]
    fn test_parse_cell_value_shared_string() {
        let parser = WorkbookParser::new();
        let strings = vec!["Hello".to_string(), "World".to_string()];
        let value = parser.parse_cell_value("0", "", Some("s"), &strings);
        assert_eq!(value, CellValue::String("Hello".to_string()));
    }

    #[test]
    fn test_parse_cell_value_boolean() {
        let parser = WorkbookParser::new();
        let value = parser.parse_cell_value("1", "", Some("b"), &[]);
        assert_eq!(value, CellValue::Boolean(true));

        let value = parser.parse_cell_value("0", "", Some("b"), &[]);
        assert_eq!(value, CellValue::Boolean(false));
    }

    #[test]
    fn test_parse_cell_value_formula() {
        let parser = WorkbookParser::new();
        let value = parser.parse_cell_value("100", "SUM(A1:A10)", None, &[]);
        match value {
            CellValue::Formula { formula, .. } => {
                assert_eq!(formula, "SUM(A1:A10)");
            }
            _ => panic!("Expected formula"),
        }
    }

    #[test]
    fn test_parse_cell_value_error() {
        let parser = WorkbookParser::new();
        let value = parser.parse_cell_value("#VALUE!", "", Some("e"), &[]);
        assert_eq!(value, CellValue::Error(CellError::Value));
    }

    #[test]
    fn test_default_trait() {
        let _parser = WorkbookParser::default();
    }

    #[test]
    fn test_parse_cell_value_boolean_true_string() {
        let parser = WorkbookParser::new();
        let value = parser.parse_cell_value("true", "", Some("b"), &[]);
        assert_eq!(value, CellValue::Boolean(true));
    }

    #[test]
    fn test_parse_cell_value_boolean_false_string() {
        let parser = WorkbookParser::new();
        let value = parser.parse_cell_value("false", "", Some("b"), &[]);
        assert_eq!(value, CellValue::Boolean(false));
    }

    #[test]
    fn test_parse_cell_value_empty() {
        let parser = WorkbookParser::new();
        let value = parser.parse_cell_value("", "", None, &[]);
        assert_eq!(value, CellValue::Empty);
    }

    #[test]
    fn test_parse_cell_value_inline_string() {
        let parser = WorkbookParser::new();
        let value = parser.parse_cell_value("Inline Text", "", Some("str"), &[]);
        assert_eq!(value, CellValue::String("Inline Text".to_string()));

        let value = parser.parse_cell_value("Another Inline", "", Some("inlineStr"), &[]);
        assert_eq!(value, CellValue::String("Another Inline".to_string()));
    }

    #[test]
    fn test_parse_cell_value_shared_string_second() {
        let parser = WorkbookParser::new();
        let strings = vec![
            "First".to_string(),
            "Second".to_string(),
            "Third".to_string(),
        ];
        let value = parser.parse_cell_value("1", "", Some("s"), &strings);
        assert_eq!(value, CellValue::String("Second".to_string()));
    }

    #[test]
    fn test_parse_cell_value_shared_string_invalid_index() {
        let parser = WorkbookParser::new();
        let strings = vec!["Only".to_string()];
        let value = parser.parse_cell_value("99", "", Some("s"), &strings);
        assert_eq!(value, CellValue::Empty);
    }

    #[test]
    fn test_parse_cell_value_shared_string_non_numeric() {
        let parser = WorkbookParser::new();
        let strings = vec!["Test".to_string()];
        let value = parser.parse_cell_value("not_a_number", "", Some("s"), &strings);
        assert_eq!(value, CellValue::Empty);
    }

    #[test]
    fn test_parse_cell_value_integer() {
        let parser = WorkbookParser::new();
        let value = parser.parse_cell_value("100", "", None, &[]);
        assert_eq!(value, CellValue::Number(100.0));
    }

    #[test]
    fn test_parse_cell_value_negative_number() {
        let parser = WorkbookParser::new();
        let value = parser.parse_cell_value("-42.5", "", None, &[]);
        assert_eq!(value, CellValue::Number(-42.5));
    }

    #[test]
    fn test_parse_cell_value_scientific_notation() {
        let parser = WorkbookParser::new();
        let value = parser.parse_cell_value("1.5e10", "", None, &[]);
        assert_eq!(value, CellValue::Number(1.5e10));
    }

    #[test]
    fn test_parse_cell_value_non_numeric_string() {
        let parser = WorkbookParser::new();
        let value = parser.parse_cell_value("Hello World", "", None, &[]);
        assert_eq!(value, CellValue::String("Hello World".to_string()));
    }

    #[test]
    fn test_parse_cell_value_formula_with_cached_result() {
        let parser = WorkbookParser::new();
        let value = parser.parse_cell_value("42", "A1+B1", None, &[]);
        match value {
            CellValue::Formula {
                formula,
                cached_result,
            } => {
                assert_eq!(formula, "A1+B1");
                assert!(cached_result.is_some());
            }
            _ => panic!("Expected formula"),
        }
    }

    #[test]
    fn test_parse_cell_value_formula_without_cached_result() {
        let parser = WorkbookParser::new();
        let value = parser.parse_cell_value("", "A1+B1", None, &[]);
        match value {
            CellValue::Formula {
                formula,
                cached_result,
            } => {
                assert_eq!(formula, "A1+B1");
                assert!(cached_result.is_none());
            }
            _ => panic!("Expected formula"),
        }
    }

    #[test]
    fn test_parse_cell_value_error_div_zero() {
        let parser = WorkbookParser::new();
        let value = parser.parse_cell_value("#DIV/0!", "", Some("e"), &[]);
        assert_eq!(value, CellValue::Error(CellError::DivZero));
    }

    #[test]
    fn test_parse_cell_value_error_ref() {
        let parser = WorkbookParser::new();
        let value = parser.parse_cell_value("#REF!", "", Some("e"), &[]);
        assert_eq!(value, CellValue::Error(CellError::Ref));
    }

    #[test]
    fn test_parse_cell_value_error_name() {
        let parser = WorkbookParser::new();
        let value = parser.parse_cell_value("#NAME?", "", Some("e"), &[]);
        assert_eq!(value, CellValue::Error(CellError::Name));
    }

    #[test]
    fn test_parse_cell_value_error_na() {
        let parser = WorkbookParser::new();
        let value = parser.parse_cell_value("#N/A", "", Some("e"), &[]);
        assert_eq!(value, CellValue::Error(CellError::Na));
    }

    #[test]
    fn test_parse_cell_value_error_null() {
        let parser = WorkbookParser::new();
        let value = parser.parse_cell_value("#NULL!", "", Some("e"), &[]);
        assert_eq!(value, CellValue::Error(CellError::Null));
    }

    #[test]
    fn test_parse_cell_value_error_num() {
        let parser = WorkbookParser::new();
        let value = parser.parse_cell_value("#NUM!", "", Some("e"), &[]);
        assert_eq!(value, CellValue::Error(CellError::Num));
    }

    #[test]
    fn test_parse_cell_value_error_unknown() {
        let parser = WorkbookParser::new();
        let value = parser.parse_cell_value("#UNKNOWN!", "", Some("e"), &[]);
        // Unknown error should return Empty based on parse implementation
        assert_eq!(value, CellValue::Empty);
    }
}
