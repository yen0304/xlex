//! styles.xml parser.

use std::io::{BufRead, Read};

use quick_xml::events::Event;
use quick_xml::Reader;

use crate::error::{XlexError, XlexResult};
use crate::style::{Border, BorderStyle, Fill, FillPattern, Font, StyleRegistry};

/// Parser for styles.xml.
pub struct StylesParser;

impl StylesParser {
    /// Creates a new parser.
    pub fn new() -> Self {
        Self
    }

    /// Parses styles from a reader.
    pub fn parse<R: Read + BufRead>(&self, reader: R) -> XlexResult<StyleRegistry> {
        let mut xml_reader = Reader::from_reader(reader);
        xml_reader.config_mut().trim_text(true);

        let mut registry = StyleRegistry::new();
        let mut buf = Vec::new();

        // Temporary storage
        let mut fonts: Vec<Font> = Vec::new();
        let mut fills: Vec<Fill> = Vec::new();
        let mut borders: Vec<Border> = Vec::new();
        let mut num_fmts: std::collections::HashMap<u32, String> = std::collections::HashMap::new();

        // Current parsing state
        let mut current_font: Option<Font> = None;
        let mut current_fill: Option<Fill> = None;
        let mut current_border: Option<Border> = None;

        loop {
            match xml_reader.read_event_into(&mut buf) {
                Ok(Event::Start(e)) | Ok(Event::Empty(e)) => match e.name().as_ref() {
                    b"font" => {
                        current_font = Some(Font::default());
                    }
                    b"name" if current_font.is_some() => {
                        if let Some(ref mut font) = current_font {
                            for attr in e.attributes().flatten() {
                                if attr.key.as_ref() == b"val" {
                                    font.name =
                                        Some(String::from_utf8_lossy(&attr.value).to_string());
                                }
                            }
                        }
                    }
                    b"sz" if current_font.is_some() => {
                        if let Some(ref mut font) = current_font {
                            for attr in e.attributes().flatten() {
                                if attr.key.as_ref() == b"val" {
                                    font.size = String::from_utf8_lossy(&attr.value).parse().ok();
                                }
                            }
                        }
                    }
                    b"b" if current_font.is_some() => {
                        if let Some(ref mut font) = current_font {
                            font.bold = true;
                        }
                    }
                    b"i" if current_font.is_some() => {
                        if let Some(ref mut font) = current_font {
                            font.italic = true;
                        }
                    }
                    b"u" if current_font.is_some() => {
                        if let Some(ref mut font) = current_font {
                            font.underline = true;
                        }
                    }
                    b"strike" if current_font.is_some() => {
                        if let Some(ref mut font) = current_font {
                            font.strikethrough = true;
                        }
                    }
                    b"fill" => {
                        current_fill = Some(Fill::default());
                    }
                    b"patternFill" if current_fill.is_some() => {
                        if let Some(ref mut fill) = current_fill {
                            for attr in e.attributes().flatten() {
                                if attr.key.as_ref() == b"patternType" {
                                    fill.pattern =
                                        parse_fill_pattern(&String::from_utf8_lossy(&attr.value));
                                }
                            }
                        }
                    }
                    b"border" => {
                        current_border = Some(Border::default());
                    }
                    b"numFmt" => {
                        let mut id: Option<u32> = None;
                        let mut code: Option<String> = None;
                        for attr in e.attributes().flatten() {
                            match attr.key.as_ref() {
                                b"numFmtId" => {
                                    id = String::from_utf8_lossy(&attr.value).parse().ok();
                                }
                                b"formatCode" => {
                                    code = Some(String::from_utf8_lossy(&attr.value).to_string());
                                }
                                _ => {}
                            }
                        }
                        if let (Some(id), Some(code)) = (id, code) {
                            num_fmts.insert(id, code);
                        }
                    }
                    _ => {}
                },
                Ok(Event::End(e)) => match e.name().as_ref() {
                    b"font" => {
                        if let Some(font) = current_font.take() {
                            fonts.push(font);
                        }
                    }
                    b"fill" => {
                        if let Some(fill) = current_fill.take() {
                            fills.push(fill);
                        }
                    }
                    b"border" => {
                        if let Some(border) = current_border.take() {
                            borders.push(border);
                        }
                    }
                    _ => {}
                },
                Ok(Event::Eof) => break,
                Err(e) => {
                    return Err(XlexError::InvalidXml {
                        message: format!("Error parsing styles: {}", e),
                    });
                }
                _ => {}
            }
            buf.clear();
        }

        // Add parsed components to registry
        for font in fonts {
            registry.add_font(font);
        }
        for fill in fills {
            registry.add_fill(fill);
        }
        for border in borders {
            registry.add_border(border);
        }
        for (id, code) in num_fmts {
            registry.add_number_format(id, code);
        }

        Ok(registry)
    }
}

impl Default for StylesParser {
    fn default() -> Self {
        Self::new()
    }
}

fn parse_fill_pattern(s: &str) -> FillPattern {
    match s {
        "none" => FillPattern::None,
        "solid" => FillPattern::Solid,
        "mediumGray" => FillPattern::MediumGray,
        "darkGray" => FillPattern::DarkGray,
        "lightGray" => FillPattern::LightGray,
        "darkHorizontal" => FillPattern::DarkHorizontal,
        "darkVertical" => FillPattern::DarkVertical,
        "darkDown" => FillPattern::DarkDown,
        "darkUp" => FillPattern::DarkUp,
        "darkGrid" => FillPattern::DarkGrid,
        "darkTrellis" => FillPattern::DarkTrellis,
        "lightHorizontal" => FillPattern::LightHorizontal,
        "lightVertical" => FillPattern::LightVertical,
        "lightDown" => FillPattern::LightDown,
        "lightUp" => FillPattern::LightUp,
        "lightGrid" => FillPattern::LightGrid,
        "lightTrellis" => FillPattern::LightTrellis,
        "gray125" => FillPattern::Gray125,
        "gray0625" => FillPattern::Gray0625,
        _ => FillPattern::None,
    }
}

#[allow(dead_code)]
fn parse_border_style(s: &str) -> BorderStyle {
    match s {
        "none" => BorderStyle::None,
        "thin" => BorderStyle::Thin,
        "medium" => BorderStyle::Medium,
        "thick" => BorderStyle::Thick,
        "dashed" => BorderStyle::Dashed,
        "dotted" => BorderStyle::Dotted,
        "double" => BorderStyle::Double,
        "hair" => BorderStyle::Hair,
        "mediumDashed" => BorderStyle::MediumDashed,
        "dashDot" => BorderStyle::DashDot,
        "mediumDashDot" => BorderStyle::MediumDashDot,
        "dashDotDot" => BorderStyle::DashDotDot,
        "mediumDashDotDot" => BorderStyle::MediumDashDotDot,
        "slantDashDot" => BorderStyle::SlantDashDot,
        _ => BorderStyle::None,
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::io::Cursor;

    #[test]
    fn test_parse_simple_styles() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8"?>
        <styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
            <fonts count="1">
                <font>
                    <sz val="11"/>
                    <name val="Calibri"/>
                </font>
            </fonts>
            <fills count="1">
                <fill>
                    <patternFill patternType="none"/>
                </fill>
            </fills>
            <borders count="1">
                <border>
                    <left/><right/><top/><bottom/><diagonal/>
                </border>
            </borders>
        </styleSheet>"#;

        let parser = StylesParser::new();
        let registry = parser.parse(Cursor::new(xml)).unwrap();

        assert_eq!(registry.fonts().len(), 1);
        assert_eq!(registry.fills().len(), 1);
        assert_eq!(registry.borders().len(), 1);
    }

    #[test]
    fn test_parse_font_styles() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8"?>
        <styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
            <fonts count="1">
                <font>
                    <b/>
                    <i/>
                    <u/>
                    <sz val="14"/>
                    <name val="Arial"/>
                </font>
            </fonts>
        </styleSheet>"#;

        let parser = StylesParser::new();
        let registry = parser.parse(Cursor::new(xml)).unwrap();

        let fonts = registry.fonts();
        assert_eq!(fonts.len(), 1);
        assert!(fonts[0].bold);
        assert!(fonts[0].italic);
        assert!(fonts[0].underline);
        assert_eq!(fonts[0].size, Some(14.0));
        assert_eq!(fonts[0].name, Some("Arial".to_string()));
    }
}
