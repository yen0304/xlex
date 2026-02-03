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

        // Write xl/styles.xml and get style ID mapping
        let style_id_map = self.write_styles(&mut zip, workbook, options)?;

        // Write xl/sharedStrings.xml if needed
        if !workbook.shared_strings().is_empty() {
            self.write_shared_strings(&mut zip, workbook, options)?;
        }

        // Write sheets
        for (index, sheet_name) in workbook.sheet_names().iter().enumerate() {
            self.write_sheet(
                &mut zip,
                workbook,
                sheet_name,
                index + 1,
                options,
                &style_id_map,
            )?;
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

        let mut content = String::from(
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
    <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
    <Default Extension="xml" ContentType="application/xml"/>
    <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
    <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
    <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
    <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
"#,
        );

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
        let mut content = String::from(
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
"#,
        );

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

        let mut content = String::from(
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
"#,
        );

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

        let mut content = String::from(
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
    <sheets>
"#,
        );

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
        workbook: &Workbook,
        options: SimpleFileOptions,
    ) -> XlexResult<std::collections::HashMap<u32, u32>> {
        use crate::style::{
            Border, BorderStyle, Fill, FillPattern, Font, HorizontalAlignment, VerticalAlignment,
        };

        zip.start_file("xl/styles.xml", options)?;

        let registry = workbook.style_registry();

        // Map from registry style ID to cellXfs index
        let mut style_id_map: std::collections::HashMap<u32, u32> =
            std::collections::HashMap::new();

        // Collect unique components from registry
        let mut fonts: Vec<Font> = vec![Font::default()]; // Default font at index 0
        let mut fills: Vec<Fill> = vec![
            Fill::default(), // none pattern at index 0
            Fill {
                pattern: FillPattern::Gray125,
                ..Default::default()
            }, // gray125 at index 1
        ];
        let mut borders: Vec<Border> = vec![Border::default()]; // Empty border at index 0
        let mut num_fmts: Vec<(u32, String)> = vec![];

        // CellXf entries: (fontId, fillId, borderId, numFmtId, style)
        struct CellXf {
            font_id: usize,
            fill_id: usize,
            border_id: usize,
            num_fmt_id: u32,
            alignment: Option<(HorizontalAlignment, VerticalAlignment, bool)>, // (h_align, v_align, wrap)
        }

        let mut cell_xfs: Vec<CellXf> = vec![CellXf {
            font_id: 0,
            fill_id: 0,
            border_id: 0,
            num_fmt_id: 0,
            alignment: None,
        }];

        // Helper to find or add font
        fn find_or_add_font(fonts: &mut Vec<Font>, font: &Font) -> usize {
            fonts.iter().position(|f| f == font).unwrap_or_else(|| {
                fonts.push(font.clone());
                fonts.len() - 1
            })
        }

        // Helper to find or add fill
        fn find_or_add_fill(fills: &mut Vec<Fill>, fill: &Fill) -> usize {
            fills.iter().position(|f| f == fill).unwrap_or_else(|| {
                fills.push(fill.clone());
                fills.len() - 1
            })
        }

        // Helper to find or add border
        fn find_or_add_border(borders: &mut Vec<Border>, border: &Border) -> usize {
            borders.iter().position(|b| b == border).unwrap_or_else(|| {
                borders.push(border.clone());
                borders.len() - 1
            })
        }

        // Process each style in registry and build mapping
        let mut next_custom_fmt_id = 164u32;
        for (style_id, style) in registry.iter() {
            let font_id = find_or_add_font(&mut fonts, &style.font);
            let fill_id = find_or_add_fill(&mut fills, &style.fill);
            let border_id = find_or_add_border(&mut borders, &style.border);

            // Handle number format
            let num_fmt_id = if let Some(code) = &style.number_format.code {
                // Custom format
                let existing = num_fmts.iter().find(|(_, c)| c == code);
                if let Some((id, _)) = existing {
                    *id
                } else {
                    let id = next_custom_fmt_id;
                    next_custom_fmt_id += 1;
                    num_fmts.push((id, code.clone()));
                    id
                }
            } else {
                style.number_format.id.unwrap_or(0)
            };

            let alignment = if style.horizontal_alignment != HorizontalAlignment::General
                || style.vertical_alignment != VerticalAlignment::Center
                || style.wrap_text
            {
                Some((
                    style.horizontal_alignment,
                    style.vertical_alignment,
                    style.wrap_text,
                ))
            } else {
                None
            };

            // Map registry style ID to cellXfs index (current length before pushing)
            let xf_index = cell_xfs.len() as u32;
            style_id_map.insert(style_id, xf_index);

            cell_xfs.push(CellXf {
                font_id,
                fill_id,
                border_id,
                num_fmt_id,
                alignment,
            });
        }

        // Generate XML
        let mut content = String::from(
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">"#,
        );

        // Number formats (if any custom)
        if !num_fmts.is_empty() {
            content.push_str(&format!(
                r#"
    <numFmts count="{}">"#,
                num_fmts.len()
            ));
            for (id, code) in &num_fmts {
                content.push_str(&format!(
                    r#"
        <numFmt numFmtId="{}" formatCode="{}"/>"#,
                    id,
                    escape_xml(code)
                ));
            }
            content.push_str(
                r#"
    </numFmts>"#,
            );
        }

        // Fonts
        content.push_str(&format!(
            r#"
    <fonts count="{}">"#,
            fonts.len()
        ));
        for font in &fonts {
            content.push_str(
                r#"
        <font>"#,
            );
            if font.bold {
                content.push_str(r#"<b/>"#);
            }
            if font.italic {
                content.push_str(r#"<i/>"#);
            }
            if font.underline {
                content.push_str(r#"<u/>"#);
            }
            if font.strikethrough {
                content.push_str(r#"<strike/>"#);
            }
            content.push_str(&format!(r#"<sz val="{}"/>"#, font.size.unwrap_or(11.0)));
            if let Some(ref color) = font.color {
                if let Some(argb) = color.to_argb_hex() {
                    content.push_str(&format!(r#"<color rgb="{}"/>"#, argb));
                }
            } else {
                content.push_str(r#"<color theme="1"/>"#);
            }
            content.push_str(&format!(
                r#"<name val="{}"/>"#,
                font.name.as_deref().unwrap_or("Calibri")
            ));
            content.push_str(r#"<family val="2"/>"#);
            content.push_str(
                r#"
        </font>"#,
            );
        }
        content.push_str(
            r#"
    </fonts>"#,
        );

        // Fills
        content.push_str(&format!(
            r#"
    <fills count="{}">"#,
            fills.len()
        ));
        for fill in &fills {
            content.push_str(
                r#"
        <fill>"#,
            );
            let pattern_type = match fill.pattern {
                FillPattern::None => "none",
                FillPattern::Solid => "solid",
                FillPattern::Gray125 => "gray125",
                FillPattern::MediumGray => "mediumGray",
                FillPattern::DarkGray => "darkGray",
                FillPattern::LightGray => "lightGray",
                _ => "solid",
            };
            if fill.pattern == FillPattern::None || fill.pattern == FillPattern::Gray125 {
                content.push_str(&format!(r#"<patternFill patternType="{}"/>"#, pattern_type));
            } else {
                content.push_str(&format!(r#"<patternFill patternType="{}">"#, pattern_type));
                if let Some(ref fg) = fill.fg_color {
                    if let Some(argb) = fg.to_argb_hex() {
                        content.push_str(&format!(r#"<fgColor rgb="{}"/>"#, argb));
                    }
                }
                if let Some(ref bg) = fill.bg_color {
                    if let Some(argb) = bg.to_argb_hex() {
                        content.push_str(&format!(r#"<bgColor rgb="{}"/>"#, argb));
                    }
                }
                content.push_str(r#"</patternFill>"#);
            }
            content.push_str(r#"</fill>"#);
        }
        content.push_str(
            r#"
    </fills>"#,
        );

        // Borders
        content.push_str(&format!(
            r#"
    <borders count="{}">"#,
            borders.len()
        ));
        for border in &borders {
            content.push_str(
                r#"
        <border>"#,
            );

            fn write_border_side(name: &str, side: &crate::style::BorderSide) -> String {
                if side.style == BorderStyle::None {
                    return format!("<{}/>", name);
                }
                let style_str = match side.style {
                    BorderStyle::Thin => "thin",
                    BorderStyle::Medium => "medium",
                    BorderStyle::Thick => "thick",
                    BorderStyle::Dashed => "dashed",
                    BorderStyle::Dotted => "dotted",
                    BorderStyle::Double => "double",
                    BorderStyle::Hair => "hair",
                    _ => "thin",
                };
                if let Some(ref color) = side.color {
                    if let Some(argb) = color.to_argb_hex() {
                        return format!(
                            r#"<{} style="{}"><color rgb="{}"/></{}>"#,
                            name, style_str, argb, name
                        );
                    }
                }
                format!(r#"<{} style="{}"/>"#, name, style_str)
            }

            content.push_str(&write_border_side("left", &border.left));
            content.push_str(&write_border_side("right", &border.right));
            content.push_str(&write_border_side("top", &border.top));
            content.push_str(&write_border_side("bottom", &border.bottom));
            content.push_str(r#"<diagonal/>"#);
            content.push_str(r#"</border>"#);
        }
        content.push_str(
            r#"
    </borders>"#,
        );

        // Cell style xfs (base styles)
        content.push_str(
            r#"
    <cellStyleXfs count="1">
        <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
    </cellStyleXfs>"#,
        );

        // Cell xfs (actual cell formats)
        content.push_str(&format!(
            r#"
    <cellXfs count="{}">"#,
            cell_xfs.len()
        ));
        for xf in &cell_xfs {
            let apply_font = if xf.font_id > 0 {
                r#" applyFont="1""#
            } else {
                ""
            };
            let apply_fill = if xf.fill_id > 0 {
                r#" applyFill="1""#
            } else {
                ""
            };
            let apply_border = if xf.border_id > 0 {
                r#" applyBorder="1""#
            } else {
                ""
            };
            let apply_fmt = if xf.num_fmt_id > 0 {
                r#" applyNumberFormat="1""#
            } else {
                ""
            };

            if let Some((h_align, v_align, wrap)) = &xf.alignment {
                let h_str = match h_align {
                    HorizontalAlignment::Left => "left",
                    HorizontalAlignment::Center => "center",
                    HorizontalAlignment::Right => "right",
                    HorizontalAlignment::Justify => "justify",
                    _ => "general",
                };
                let v_str = match v_align {
                    VerticalAlignment::Top => "top",
                    VerticalAlignment::Center => "center",
                    VerticalAlignment::Bottom => "bottom",
                    _ => "center",
                };
                let wrap_attr = if *wrap { r#" wrapText="1""# } else { "" };
                content.push_str(&format!(
                    r#"
        <xf numFmtId="{}" fontId="{}" fillId="{}" borderId="{}" xfId="0"{}{}{}{} applyAlignment="1"><alignment horizontal="{}" vertical="{}"{}/></xf>"#,
                    xf.num_fmt_id, xf.font_id, xf.fill_id, xf.border_id,
                    apply_font, apply_fill, apply_border, apply_fmt,
                    h_str, v_str, wrap_attr
                ));
            } else {
                content.push_str(&format!(
                    r#"
        <xf numFmtId="{}" fontId="{}" fillId="{}" borderId="{}" xfId="0"{}{}{}{}/>"#,
                    xf.num_fmt_id,
                    xf.font_id,
                    xf.fill_id,
                    xf.border_id,
                    apply_font,
                    apply_fill,
                    apply_border,
                    apply_fmt
                ));
            }
        }
        content.push_str(
            r#"
    </cellXfs>"#,
        );

        content.push_str(
            r#"
</styleSheet>"#,
        );

        zip.write_all(content.as_bytes())?;
        Ok(style_id_map)
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
        style_id_map: &std::collections::HashMap<u32, u32>,
    ) -> XlexResult<()> {
        let sheet = workbook
            .get_sheet(sheet_name)
            .ok_or_else(|| XlexError::SheetNotFound {
                name: sheet_name.to_string(),
            })?;

        zip.start_file(format!("xl/worksheets/sheet{}.xml", sheet_number), options)?;

        let mut content = String::from(
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
    <sheetData>
"#,
        );

        // Collect cells by row
        let mut rows: std::collections::BTreeMap<u32, Vec<&crate::cell::Cell>> =
            std::collections::BTreeMap::new();

        for cell in sheet.cells() {
            rows.entry(cell.reference.row).or_default().push(cell);
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

                let type_attr = cell_type
                    .map(|t| format!(r#" t="{}""#, t))
                    .unwrap_or_default();

                // Map cell's style_id to cellXfs index using the mapping
                let style_attr = cell
                    .style_id
                    .and_then(|registry_id| style_id_map.get(&registry_id))
                    .map(|xf_index| format!(r#" s="{}""#, xf_index))
                    .unwrap_or_default();

                match &cell.value {
                    CellValue::Formula { formula, .. } => {
                        content.push_str(&format!(
                            r#"<c r="{}"{}{}><f>{}</f>{}</c>"#,
                            cell_ref,
                            type_attr,
                            style_attr,
                            escape_xml(formula),
                            cell_value
                                .map(|v| format!("<v>{}</v>", v))
                                .unwrap_or_default()
                        ));
                    }
                    CellValue::Empty => {
                        // Skip empty cells
                    }
                    CellValue::String(_) => {
                        // inlineStr type requires <is><t>...</t></is> format
                        if let Some(value) = cell_value {
                            content.push_str(&format!(
                                r#"<c r="{}"{}{}><is><t>{}</t></is></c>"#,
                                cell_ref, type_attr, style_attr, value
                            ));
                        }
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

        content.push_str("    </sheetData>\n");

        // Write merged cells if any
        let merged_ranges = sheet.merged_ranges();
        if !merged_ranges.is_empty() {
            content.push_str(&format!(
                r#"    <mergeCells count="{}">"#,
                merged_ranges.len()
            ));
            for range in merged_ranges {
                content.push_str(&format!(r#"<mergeCell ref="{}"/>"#, range.to_a1()));
            }
            content.push_str("</mergeCells>\n");
        }

        content.push_str("</worksheet>");

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
    use crate::cell::CellError;

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

    #[test]
    fn test_default_trait() {
        let _writer = WorkbookWriter::default();
    }

    #[test]
    fn test_escape_xml_single_quote() {
        assert_eq!(escape_xml("it's"), "it&apos;s");
    }

    #[test]
    fn test_escape_xml_combined() {
        let input = r#"<tag attr="val's" & more>"#;
        let expected = "&lt;tag attr=&quot;val&apos;s&quot; &amp; more&gt;";
        assert_eq!(escape_xml(input), expected);
    }

    #[test]
    fn test_escape_xml_empty() {
        assert_eq!(escape_xml(""), "");
    }

    #[test]
    fn test_escape_xml_no_special_chars() {
        assert_eq!(escape_xml("Normal text 123"), "Normal text 123");
    }

    #[test]
    fn test_format_cell_value_empty() {
        let writer = WorkbookWriter::new();
        let (t, v) = writer.format_cell_value(&CellValue::Empty);
        assert!(t.is_none());
        assert!(v.is_none());
    }

    #[test]
    fn test_format_cell_value_string() {
        let writer = WorkbookWriter::new();
        let (t, v) = writer.format_cell_value(&CellValue::String("Hello".to_string()));
        assert_eq!(t, Some("inlineStr"));
        assert_eq!(v, Some("Hello".to_string()));
    }

    #[test]
    fn test_format_cell_value_string_with_xml() {
        let writer = WorkbookWriter::new();
        let (t, v) = writer.format_cell_value(&CellValue::String("<test>".to_string()));
        assert_eq!(t, Some("inlineStr"));
        assert_eq!(v, Some("&lt;test&gt;".to_string()));
    }

    #[test]
    fn test_format_cell_value_boolean_false() {
        let writer = WorkbookWriter::new();
        let (t, v) = writer.format_cell_value(&CellValue::Boolean(false));
        assert_eq!(t, Some("b"));
        assert_eq!(v, Some("0".to_string()));
    }

    #[test]
    fn test_format_cell_value_number_decimal() {
        let writer = WorkbookWriter::new();
        let (t, v) = writer.format_cell_value(&CellValue::Number(3.14159));
        assert!(t.is_none());
        assert_eq!(v, Some("3.14159".to_string()));
    }

    #[test]
    fn test_format_cell_value_number_integer() {
        let writer = WorkbookWriter::new();
        let (t, v) = writer.format_cell_value(&CellValue::Number(100.0));
        assert!(t.is_none());
        assert_eq!(v, Some("100".to_string()));
    }

    #[test]
    fn test_format_cell_value_number_negative() {
        let writer = WorkbookWriter::new();
        let (t, v) = writer.format_cell_value(&CellValue::Number(-42.5));
        assert!(t.is_none());
        assert_eq!(v, Some("-42.5".to_string()));
    }

    #[test]
    fn test_format_cell_value_number_large() {
        let writer = WorkbookWriter::new();
        let (t, v) = writer.format_cell_value(&CellValue::Number(1e16));
        assert!(t.is_none());
        assert!(v.is_some());
    }

    #[test]
    fn test_format_cell_value_formula_with_cached() {
        let writer = WorkbookWriter::new();
        let value = CellValue::Formula {
            formula: "SUM(A1:A10)".to_string(),
            cached_result: Some(Box::new(CellValue::Number(100.0))),
        };
        let (t, v) = writer.format_cell_value(&value);
        assert!(t.is_none());
        assert_eq!(v, Some("100".to_string()));
    }

    #[test]
    fn test_format_cell_value_formula_without_cached() {
        let writer = WorkbookWriter::new();
        let value = CellValue::Formula {
            formula: "SUM(A1:A10)".to_string(),
            cached_result: None,
        };
        let (t, v) = writer.format_cell_value(&value);
        assert!(t.is_none());
        assert!(v.is_none());
    }

    #[test]
    fn test_format_cell_value_error() {
        let writer = WorkbookWriter::new();
        let (t, v) = writer.format_cell_value(&CellValue::Error(CellError::Value));
        assert_eq!(t, Some("e"));
        assert_eq!(v, Some("#VALUE!".to_string()));
    }

    #[test]
    fn test_format_cell_value_error_div_zero() {
        let writer = WorkbookWriter::new();
        let (t, v) = writer.format_cell_value(&CellValue::Error(CellError::DivZero));
        assert_eq!(t, Some("e"));
        assert_eq!(v, Some("#DIV/0!".to_string()));
    }

    #[test]
    fn test_format_cell_value_datetime() {
        let writer = WorkbookWriter::new();
        let (t, v) = writer.format_cell_value(&CellValue::DateTime(44927.5));
        assert!(t.is_none());
        assert_eq!(v, Some("44927.5".to_string()));
    }

    #[test]
    fn test_format_cell_value_formula_cached_string() {
        let writer = WorkbookWriter::new();
        let value = CellValue::Formula {
            formula: "CONCAT(A1,B1)".to_string(),
            cached_result: Some(Box::new(CellValue::String("HelloWorld".to_string()))),
        };
        let (t, v) = writer.format_cell_value(&value);
        // String type from cached result
        assert!(t.is_none());
        assert_eq!(v, Some("HelloWorld".to_string()));
    }

    #[test]
    fn test_format_cell_value_formula_cached_boolean() {
        let writer = WorkbookWriter::new();
        let value = CellValue::Formula {
            formula: "A1>B1".to_string(),
            cached_result: Some(Box::new(CellValue::Boolean(true))),
        };
        let (t, v) = writer.format_cell_value(&value);
        assert!(t.is_none());
        assert_eq!(v, Some("1".to_string()));
    }

    #[test]
    fn test_escape_xml_unicode() {
        assert_eq!(escape_xml("Hello 世界"), "Hello 世界");
        assert_eq!(escape_xml("🎉 Party"), "🎉 Party");
    }

    #[test]
    fn test_escape_xml_multiple_ampersands() {
        assert_eq!(escape_xml("A & B & C"), "A &amp; B &amp; C");
    }

    #[test]
    fn test_escape_xml_nested_tags() {
        let input = "<outer><inner/></outer>";
        let expected = "&lt;outer&gt;&lt;inner/&gt;&lt;/outer&gt;";
        assert_eq!(escape_xml(input), expected);
    }

    #[test]
    fn test_workbook_writer_new() {
        let writer = WorkbookWriter::new();
        // Just ensure it can be created
        let _ = writer;
    }
}
