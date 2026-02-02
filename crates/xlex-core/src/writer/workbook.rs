//! Workbook writer.

use std::fs::File;
use std::io::{BufWriter, Write};
use std::path::Path;

use zip::write::SimpleFileOptions;
use zip::ZipWriter;

use crate::cell::CellValue;
use crate::error::{XlexError, XlexResult};
use crate::workbook::Workbook;

/// Writer for xlsx workbooks.
pub struct WorkbookWriter;

impl WorkbookWriter {
    /// Creates a new workbook writer.
    pub fn new() -> Self {
        Self
    }

    /// Writes a workbook to a file.
    pub fn write(&self, workbook: &Workbook, path: &Path) -> XlexResult<()> {
        // Create temp file
        let temp_path = path.with_extension("xlsx.tmp");
        let file = File::create(&temp_path)?;
        let writer = BufWriter::new(file);

        let result = self.write_to_zip(workbook, writer);

        if result.is_ok() {
            // Atomic rename
            std::fs::rename(&temp_path, path)?;
        } else {
            // Cleanup temp file
            let _ = std::fs::remove_file(&temp_path);
        }

        result
    }

    /// Writes a workbook to a ZIP writer.
    fn write_to_zip<W: Write + std::io::Seek>(
        &self,
        workbook: &Workbook,
        writer: W,
    ) -> XlexResult<()> {
        let mut zip = ZipWriter::new(writer);
        let options = SimpleFileOptions::default()
            .compression_method(zip::CompressionMethod::Deflated)
            .compression_level(Some(6));

        // Write [Content_Types].xml
        self.write_content_types(&mut zip, workbook, options)?;

        // Write _rels/.rels
        self.write_root_rels(&mut zip, options)?;

        // Write docProps/app.xml
        self.write_app_props(&mut zip, options)?;

        // Write docProps/core.xml
        self.write_core_props(&mut zip, workbook, options)?;

        // Write xl/_rels/workbook.xml.rels
        self.write_workbook_rels(&mut zip, workbook, options)?;

        // Write xl/workbook.xml
        self.write_workbook_xml(&mut zip, workbook, options)?;

        // Write xl/styles.xml
        self.write_styles(&mut zip, workbook, options)?;

        // Write xl/sharedStrings.xml if needed
        if !workbook.shared_strings().is_empty() {
            self.write_shared_strings(&mut zip, workbook, options)?;
        }

        // Write sheets
        for (index, sheet_name) in workbook.sheet_names().iter().enumerate() {
            self.write_sheet(&mut zip, workbook, sheet_name, index + 1, options)?;
        }

        zip.finish()?;
        Ok(())
    }

    fn write_content_types<W: Write + std::io::Seek>(
        &self,
        zip: &mut ZipWriter<W>,
        workbook: &Workbook,
        options: SimpleFileOptions,
    ) -> XlexResult<()> {
        zip.start_file("[Content_Types].xml", options)?;

        let mut content = String::from(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
    <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
    <Default Extension="xml" ContentType="application/xml"/>
    <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
    <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
    <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
    <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
"#);

        // Add shared strings if present
        if !workbook.shared_strings().is_empty() {
            content.push_str(r#"    <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
"#);
        }

        // Add sheets
        for (index, _) in workbook.sheet_names().iter().enumerate() {
            content.push_str(&format!(
                r#"    <Override PartName="/xl/worksheets/sheet{}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
"#,
                index + 1
            ));
        }

        content.push_str("</Types>");
        zip.write_all(content.as_bytes())?;
        Ok(())
    }

    fn write_root_rels<W: Write + std::io::Seek>(
        &self,
        zip: &mut ZipWriter<W>,
        options: SimpleFileOptions,
    ) -> XlexResult<()> {
        zip.start_file("_rels/.rels", options)?;

        let content = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
    <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
    <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>"#;

        zip.write_all(content.as_bytes())?;
        Ok(())
    }

    fn write_app_props<W: Write + std::io::Seek>(
        &self,
        zip: &mut ZipWriter<W>,
        options: SimpleFileOptions,
    ) -> XlexResult<()> {
        zip.start_file("docProps/app.xml", options)?;

        let content = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">
    <Application>XLEX</Application>
</Properties>"#;

        zip.write_all(content.as_bytes())?;
        Ok(())
    }

    fn write_core_props<W: Write + std::io::Seek>(
        &self,
        zip: &mut ZipWriter<W>,
        workbook: &Workbook,
        options: SimpleFileOptions,
    ) -> XlexResult<()> {
        zip.start_file("docProps/core.xml", options)?;

        let props = workbook.properties();
        let mut content = String::from(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
"#);

        if let Some(ref title) = props.title {
            content.push_str(&format!("    <dc:title>{}</dc:title>\n", escape_xml(title)));
        }
        if let Some(ref creator) = props.creator {
            content.push_str(&format!(
                "    <dc:creator>{}</dc:creator>\n",
                escape_xml(creator)
            ));
        }
        if let Some(ref subject) = props.subject {
            content.push_str(&format!(
                "    <dc:subject>{}</dc:subject>\n",
                escape_xml(subject)
            ));
        }
        if let Some(ref description) = props.description {
            content.push_str(&format!(
                "    <dc:description>{}</dc:description>\n",
                escape_xml(description)
            ));
        }
        if let Some(ref keywords) = props.keywords {
            content.push_str(&format!(
                "    <cp:keywords>{}</cp:keywords>\n",
                escape_xml(keywords)
            ));
        }
        if let Some(ref last_modified_by) = props.last_modified_by {
            content.push_str(&format!(
                "    <cp:lastModifiedBy>{}</cp:lastModifiedBy>\n",
                escape_xml(last_modified_by)
            ));
        }

        content.push_str("</cp:coreProperties>");
        zip.write_all(content.as_bytes())?;
        Ok(())
    }

    fn write_workbook_rels<W: Write + std::io::Seek>(
        &self,
        zip: &mut ZipWriter<W>,
        workbook: &Workbook,
        options: SimpleFileOptions,
    ) -> XlexResult<()> {
        zip.start_file("xl/_rels/workbook.xml.rels", options)?;

        let mut content = String::from(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
"#);

        // Add sheet relationships
        for (index, _) in workbook.sheet_names().iter().enumerate() {
            content.push_str(&format!(
                r#"    <Relationship Id="rId{}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet{}.xml"/>
"#,
                index + 1,
                index + 1
            ));
        }

        // Styles relationship
        let styles_rid = workbook.sheet_count() + 1;
        content.push_str(&format!(
            r#"    <Relationship Id="rId{}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
"#,
            styles_rid
        ));

        // Shared strings relationship if present
        if !workbook.shared_strings().is_empty() {
            let strings_rid = styles_rid + 1;
            content.push_str(&format!(
                r#"    <Relationship Id="rId{}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>
"#,
                strings_rid
            ));
        }

        content.push_str("</Relationships>");
        zip.write_all(content.as_bytes())?;
        Ok(())
    }

    fn write_workbook_xml<W: Write + std::io::Seek>(
        &self,
        zip: &mut ZipWriter<W>,
        workbook: &Workbook,
        options: SimpleFileOptions,
    ) -> XlexResult<()> {
        zip.start_file("xl/workbook.xml", options)?;

        let mut content = String::from(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
    <sheets>
"#);

        for (index, name) in workbook.sheet_names().iter().enumerate() {
            let visibility = workbook.get_sheet_visibility(name).unwrap_or_default();
            let state_attr = match visibility {
                crate::sheet::SheetVisibility::Visible => "",
                crate::sheet::SheetVisibility::Hidden => r#" state="hidden""#,
                crate::sheet::SheetVisibility::VeryHidden => r#" state="veryHidden""#,
            };
            content.push_str(&format!(
                r#"        <sheet name="{}" sheetId="{}" r:id="rId{}"{}/>
"#,
                escape_xml(name),
                index + 1,
                index + 1,
                state_attr
            ));
        }

        content.push_str("    </sheets>\n");

        // Write defined names if any
        let defined_names = workbook.defined_names();
        if !defined_names.is_empty() {
            content.push_str("    <definedNames>\n");
            for dn in defined_names {
                let mut attrs = format!(r#"name="{}""#, escape_xml(&dn.name));
                if let Some(sheet_id) = dn.local_sheet_id {
                    attrs.push_str(&format!(r#" localSheetId="{}""#, sheet_id));
                }
                if let Some(ref comment) = dn.comment {
                    attrs.push_str(&format!(r#" comment="{}""#, escape_xml(comment)));
                }
                if dn.hidden {
                    attrs.push_str(r#" hidden="1""#);
                }
                content.push_str(&format!(
                    "        <definedName {}>{}</definedName>\n",
                    attrs,
                    escape_xml(&dn.reference)
                ));
            }
            content.push_str("    </definedNames>\n");
        }

        content.push_str("</workbook>");

        zip.write_all(content.as_bytes())?;
        Ok(())
    }

    fn write_styles<W: Write + std::io::Seek>(
        &self,
        zip: &mut ZipWriter<W>,
        _workbook: &Workbook,
        options: SimpleFileOptions,
    ) -> XlexResult<()> {
        zip.start_file("xl/styles.xml", options)?;

        // Minimal styles.xml
        let content = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
    <fonts count="1">
        <font>
            <sz val="11"/>
            <color theme="1"/>
            <name val="Calibri"/>
            <family val="2"/>
            <scheme val="minor"/>
        </font>
    </fonts>
    <fills count="2">
        <fill><patternFill patternType="none"/></fill>
        <fill><patternFill patternType="gray125"/></fill>
    </fills>
    <borders count="1">
        <border><left/><right/><top/><bottom/><diagonal/></border>
    </borders>
    <cellStyleXfs count="1">
        <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
    </cellStyleXfs>
    <cellXfs count="1">
        <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
    </cellXfs>
</styleSheet>"#;

        zip.write_all(content.as_bytes())?;
        Ok(())
    }

    fn write_shared_strings<W: Write + std::io::Seek>(
        &self,
        zip: &mut ZipWriter<W>,
        workbook: &Workbook,
        options: SimpleFileOptions,
    ) -> XlexResult<()> {
        zip.start_file("xl/sharedStrings.xml", options)?;

        let strings = workbook.shared_strings();
        let mut content = format!(
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="{}" uniqueCount="{}">
"#,
            strings.len(),
            strings.len()
        );

        for s in strings {
            content.push_str(&format!("    <si><t>{}</t></si>\n", escape_xml(s)));
        }

        content.push_str("</sst>");
        zip.write_all(content.as_bytes())?;
        Ok(())
    }

    fn write_sheet<W: Write + std::io::Seek>(
        &self,
        zip: &mut ZipWriter<W>,
        workbook: &Workbook,
        sheet_name: &str,
        sheet_number: usize,
        options: SimpleFileOptions,
    ) -> XlexResult<()> {
        let sheet = workbook.get_sheet(sheet_name).ok_or_else(|| XlexError::SheetNotFound {
            name: sheet_name.to_string(),
        })?;

        zip.start_file(format!("xl/worksheets/sheet{}.xml", sheet_number), options)?;

        let mut content = String::from(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
    <sheetData>
"#);

        // Collect cells by row
        let mut rows: std::collections::BTreeMap<u32, Vec<&crate::cell::Cell>> =
            std::collections::BTreeMap::new();

        for cell in sheet.cells() {
            rows.entry(cell.reference.row)
                .or_default()
                .push(cell);
        }

        // Write rows
        for (row_num, cells) in rows {
            content.push_str(&format!(r#"        <row r="{}">"#, row_num));

            // Sort cells by column
            let mut sorted_cells = cells;
            sorted_cells.sort_by_key(|c| c.reference.col);

            for cell in sorted_cells {
                let cell_ref = cell.reference.to_a1();
                let (cell_type, cell_value) = self.format_cell_value(&cell.value);

                let type_attr = cell_type.map(|t| format!(r#" t="{}""#, t)).unwrap_or_default();
                let style_attr = cell
                    .style_id
                    .map(|s| format!(r#" s="{}""#, s))
                    .unwrap_or_default();

                match &cell.value {
                    CellValue::Formula { formula, .. } => {
                        content.push_str(&format!(
                            r#"<c r="{}"{}{}><f>{}</f>{}</c>"#,
                            cell_ref,
                            type_attr,
                            style_attr,
                            escape_xml(formula),
                            cell_value.map(|v| format!("<v>{}</v>", v)).unwrap_or_default()
                        ));
                    }
                    CellValue::Empty => {
                        // Skip empty cells
                    }
                    _ => {
                        if let Some(value) = cell_value {
                            content.push_str(&format!(
                                r#"<c r="{}"{}{}><v>{}</v></c>"#,
                                cell_ref, type_attr, style_attr, value
                            ));
                        }
                    }
                }
            }

            content.push_str("</row>\n");
        }

        content.push_str(r#"    </sheetData>
</worksheet>"#);

        zip.write_all(content.as_bytes())?;
        Ok(())
    }

    /// Formats a cell value for XML output.
    /// Returns (type_attribute, value_string).
    fn format_cell_value(&self, value: &CellValue) -> (Option<&'static str>, Option<String>) {
        match value {
            CellValue::Empty => (None, None),
            CellValue::String(s) => (Some("inlineStr"), Some(escape_xml(s))),
            CellValue::Number(n) => {
                let formatted = if n.fract() == 0.0 && n.abs() < 1e15 {
                    format!("{:.0}", n)
                } else {
                    n.to_string()
                };
                (None, Some(formatted))
            }
            CellValue::Boolean(b) => (Some("b"), Some(if *b { "1" } else { "0" }.to_string())),
            CellValue::Formula { cached_result, .. } => {
                // Return cached result value if available
                if let Some(cached) = cached_result {
                    let (_, val) = self.format_cell_value(cached);
                    (None, val)
                } else {
                    (None, None)
                }
            }
            CellValue::Error(e) => (Some("e"), Some(e.to_string())),
            CellValue::DateTime(serial) => (None, Some(serial.to_string())),
        }
    }
}

impl Default for WorkbookWriter {
    fn default() -> Self {
        Self::new()
    }
}

/// Escapes special XML characters.
fn escape_xml(s: &str) -> String {
    s.replace('&', "&amp;")
        .replace('<', "&lt;")
        .replace('>', "&gt;")
        .replace('"', "&quot;")
        .replace('\'', "&apos;")
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_escape_xml() {
        assert_eq!(escape_xml("Hello"), "Hello");
        assert_eq!(escape_xml("<test>"), "&lt;test&gt;");
        assert_eq!(escape_xml("A & B"), "A &amp; B");
        assert_eq!(escape_xml(r#""quoted""#), "&quot;quoted&quot;");
    }

    #[test]
    fn test_format_cell_value_number() {
        let writer = WorkbookWriter::new();
        let (t, v) = writer.format_cell_value(&CellValue::Number(42.0));
        assert!(t.is_none());
        assert_eq!(v, Some("42".to_string()));
    }

    #[test]
    fn test_format_cell_value_boolean() {
        let writer = WorkbookWriter::new();
        let (t, v) = writer.format_cell_value(&CellValue::Boolean(true));
        assert_eq!(t, Some("b"));
        assert_eq!(v, Some("1".to_string()));
    }
}
