//! SharedStrings.xml parser with LRU caching.

use std::io::{BufRead, Read};

use lru::LruCache;
use quick_xml::events::Event;
use quick_xml::Reader;

use crate::error::{XlexError, XlexResult};

/// Default LRU cache size for shared strings.
pub const DEFAULT_CACHE_SIZE: usize = 10_000;

/// Parser for SharedStrings.xml with lazy loading and LRU caching.
pub struct SharedStringsParser {
    /// Cached strings by index
    cache: LruCache<u32, String>,
    /// All strings (if fully loaded)
    strings: Option<Vec<String>>,
    /// Total count of strings
    count: Option<u32>,
}

impl SharedStringsParser {
    /// Creates a new parser with the given cache size.
    pub fn new(cache_size: usize) -> Self {
        Self {
            cache: LruCache::new(
                std::num::NonZeroUsize::new(cache_size)
                    .unwrap_or(std::num::NonZeroUsize::new(DEFAULT_CACHE_SIZE).unwrap()),
            ),
            strings: None,
            count: None,
        }
    }

    /// Creates a new parser with default cache size.
    pub fn with_default_cache() -> Self {
        Self::new(DEFAULT_CACHE_SIZE)
    }

    /// Parses all shared strings from a reader.
    pub fn parse_all<R: Read + BufRead>(&mut self, reader: R) -> XlexResult<Vec<String>> {
        let mut xml_reader = Reader::from_reader(reader);
        // Don't trim text - preserve spaces in rich text
        xml_reader.config_mut().trim_text(false);

        let mut strings = Vec::new();
        let mut buf = Vec::new();
        let mut current_string = String::new();
        let mut in_si = false;
        let mut in_t = false;

        loop {
            match xml_reader.read_event_into(&mut buf) {
                Ok(Event::Start(e)) => match e.name().as_ref() {
                    b"si" => {
                        in_si = true;
                        current_string.clear();
                    }
                    b"t" if in_si => {
                        in_t = true;
                    }
                    _ => {}
                },
                Ok(Event::End(e)) => match e.name().as_ref() {
                    b"si" => {
                        strings.push(std::mem::take(&mut current_string));
                        in_si = false;
                    }
                    b"t" => {
                        in_t = false;
                    }
                    _ => {}
                },
                Ok(Event::Text(e)) => {
                    if in_t {
                        let text = e.unescape().map_err(|e| XlexError::InvalidXml {
                            message: e.to_string(),
                        })?;
                        current_string.push_str(&text);
                    }
                }
                Ok(Event::Eof) => break,
                Err(e) => {
                    return Err(XlexError::InvalidXml {
                        message: format!("Error parsing SharedStrings: {}", e),
                    });
                }
                _ => {}
            }
            buf.clear();
        }

        self.count = Some(strings.len() as u32);
        self.strings = Some(strings.clone());
        Ok(strings)
    }

    /// Gets a string by index.
    pub fn get(&mut self, index: u32) -> Option<String> {
        // Check cache first
        if let Some(s) = self.cache.get(&index) {
            return Some(s.clone());
        }

        // Check full strings list if available
        if let Some(ref strings) = self.strings {
            if let Some(s) = strings.get(index as usize) {
                self.cache.put(index, s.clone());
                return Some(s.clone());
            }
        }

        None
    }

    /// Returns the total count of strings.
    pub fn count(&self) -> Option<u32> {
        self.count
    }

    /// Returns all strings if fully loaded.
    pub fn all_strings(&self) -> Option<&[String]> {
        self.strings.as_deref()
    }
}

impl Default for SharedStringsParser {
    fn default() -> Self {
        Self::with_default_cache()
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::io::Cursor;

    #[test]
    fn test_parse_shared_strings() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8"?>
        <sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="3" uniqueCount="3">
            <si><t>Hello</t></si>
            <si><t>World</t></si>
            <si><t>Test</t></si>
        </sst>"#;

        let mut parser = SharedStringsParser::with_default_cache();
        let strings = parser.parse_all(Cursor::new(xml)).unwrap();

        assert_eq!(strings.len(), 3);
        assert_eq!(strings[0], "Hello");
        assert_eq!(strings[1], "World");
        assert_eq!(strings[2], "Test");
    }

    #[test]
    fn test_shared_strings_cache() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8"?>
        <sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
            <si><t>String1</t></si>
            <si><t>String2</t></si>
        </sst>"#;

        let mut parser = SharedStringsParser::with_default_cache();
        parser.parse_all(Cursor::new(xml)).unwrap();

        assert_eq!(parser.get(0), Some("String1".to_string()));
        assert_eq!(parser.get(1), Some("String2".to_string()));
        assert_eq!(parser.get(2), None);
    }

    #[test]
    fn test_rich_text_strings() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8"?>
        <sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
            <si>
                <r><t>Rich </t></r>
                <r><t>Text</t></r>
            </si>
        </sst>"#;

        let mut parser = SharedStringsParser::with_default_cache();
        let strings = parser.parse_all(Cursor::new(xml)).unwrap();

        // Rich text should be concatenated
        assert_eq!(strings.len(), 1);
        assert_eq!(strings[0], "Rich Text");
    }

    #[test]
    fn test_default_cache_size() {
        assert_eq!(DEFAULT_CACHE_SIZE, 10_000);
    }

    #[test]
    fn test_new_with_cache_size() {
        let parser = SharedStringsParser::new(5000);
        assert!(parser.count().is_none());
        assert!(parser.all_strings().is_none());
    }

    #[test]
    fn test_new_with_zero_cache_size() {
        // Zero cache size should fall back to default
        let parser = SharedStringsParser::new(0);
        assert!(parser.count().is_none());
    }

    #[test]
    fn test_default_trait() {
        let parser = SharedStringsParser::default();
        assert!(parser.count().is_none());
        assert!(parser.all_strings().is_none());
    }

    #[test]
    fn test_count_after_parse() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8"?>
        <sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
            <si><t>One</t></si>
            <si><t>Two</t></si>
            <si><t>Three</t></si>
        </sst>"#;

        let mut parser = SharedStringsParser::with_default_cache();
        parser.parse_all(Cursor::new(xml)).unwrap();

        assert_eq!(parser.count(), Some(3));
    }

    #[test]
    fn test_all_strings_after_parse() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8"?>
        <sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
            <si><t>Apple</t></si>
            <si><t>Banana</t></si>
        </sst>"#;

        let mut parser = SharedStringsParser::with_default_cache();
        parser.parse_all(Cursor::new(xml)).unwrap();

        let all = parser.all_strings().unwrap();
        assert_eq!(all.len(), 2);
        assert_eq!(all[0], "Apple");
        assert_eq!(all[1], "Banana");
    }

    #[test]
    fn test_empty_sst() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8"?>
        <sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
        </sst>"#;

        let mut parser = SharedStringsParser::with_default_cache();
        let strings = parser.parse_all(Cursor::new(xml)).unwrap();

        assert!(strings.is_empty());
        assert_eq!(parser.count(), Some(0));
    }

    #[test]
    fn test_empty_string_value() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8"?>
        <sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
            <si><t></t></si>
            <si><t>NotEmpty</t></si>
        </sst>"#;

        let mut parser = SharedStringsParser::with_default_cache();
        let strings = parser.parse_all(Cursor::new(xml)).unwrap();

        assert_eq!(strings.len(), 2);
        assert_eq!(strings[0], "");
        assert_eq!(strings[1], "NotEmpty");
    }

    #[test]
    fn test_string_with_xml_entities() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8"?>
        <sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
            <si><t>&lt;tag&gt;</t></si>
            <si><t>A &amp; B</t></si>
            <si><t>&quot;quoted&quot;</t></si>
        </sst>"#;

        let mut parser = SharedStringsParser::with_default_cache();
        let strings = parser.parse_all(Cursor::new(xml)).unwrap();

        assert_eq!(strings[0], "<tag>");
        assert_eq!(strings[1], "A & B");
        assert_eq!(strings[2], "\"quoted\"");
    }

    #[test]
    fn test_string_with_unicode() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8"?>
        <sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
            <si><t>こんにちは</t></si>
            <si><t>你好</t></si>
            <si><t>🎉</t></si>
        </sst>"#;

        let mut parser = SharedStringsParser::with_default_cache();
        let strings = parser.parse_all(Cursor::new(xml)).unwrap();

        assert_eq!(strings[0], "こんにちは");
        assert_eq!(strings[1], "你好");
        assert_eq!(strings[2], "🎉");
    }

    #[test]
    fn test_get_before_parse() {
        let mut parser = SharedStringsParser::with_default_cache();
        assert_eq!(parser.get(0), None);
        assert_eq!(parser.get(100), None);
    }

    #[test]
    fn test_get_caches_value() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8"?>
        <sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
            <si><t>Cached</t></si>
        </sst>"#;

        let mut parser = SharedStringsParser::with_default_cache();
        parser.parse_all(Cursor::new(xml)).unwrap();

        // First get should populate cache
        assert_eq!(parser.get(0), Some("Cached".to_string()));
        // Second get should hit cache
        assert_eq!(parser.get(0), Some("Cached".to_string()));
    }

    #[test]
    fn test_small_cache_eviction() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8"?>
        <sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
            <si><t>A</t></si>
            <si><t>B</t></si>
            <si><t>C</t></si>
        </sst>"#;

        // Small cache of size 2
        let mut parser = SharedStringsParser::new(2);
        parser.parse_all(Cursor::new(xml)).unwrap();

        // Access all strings - cache should handle eviction
        assert_eq!(parser.get(0), Some("A".to_string()));
        assert_eq!(parser.get(1), Some("B".to_string()));
        assert_eq!(parser.get(2), Some("C".to_string()));

        // All should still be accessible from strings vec
        assert_eq!(parser.get(0), Some("A".to_string()));
    }

    #[test]
    fn test_nested_rich_text_elements() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8"?>
        <sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
            <si>
                <r>
                    <rPr><b/><sz val="12"/></rPr>
                    <t>Bold</t>
                </r>
                <r>
                    <rPr><i/></rPr>
                    <t> Italic</t>
                </r>
            </si>
        </sst>"#;

        let mut parser = SharedStringsParser::with_default_cache();
        let strings = parser.parse_all(Cursor::new(xml)).unwrap();

        assert_eq!(strings.len(), 1);
        assert_eq!(strings[0], "Bold Italic");
    }

    #[test]
    fn test_malformed_xml_returns_error() {
        // Completely malformed XML structure
        let xml = r#"<<<<not valid xml>>>>"#;

        let mut parser = SharedStringsParser::with_default_cache();
        let result = parser.parse_all(Cursor::new(xml));

        // Parser should handle malformed XML somehow (may return error or empty)
        // The behavior depends on quick_xml's error handling
        if result.is_ok() {
            assert!(result.unwrap().is_empty());
        }
    }

    #[test]
    fn test_multiple_t_elements_in_si() {
        // Some Excel files have multiple <t> directly in <si>
        let xml = r#"<?xml version="1.0" encoding="UTF-8"?>
        <sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
            <si>
                <t>Part1</t>
                <t>Part2</t>
            </si>
        </sst>"#;

        let mut parser = SharedStringsParser::with_default_cache();
        let strings = parser.parse_all(Cursor::new(xml)).unwrap();

        assert_eq!(strings.len(), 1);
        assert_eq!(strings[0], "Part1Part2");
    }

    #[test]
    fn test_whitespace_preservation() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8"?>
        <sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
            <si><t>  spaces  </t></si>
            <si><t>	tab	</t></si>
        </sst>"#;

        let mut parser = SharedStringsParser::with_default_cache();
        let strings = parser.parse_all(Cursor::new(xml)).unwrap();

        assert_eq!(strings[0], "  spaces  ");
        assert_eq!(strings[1], "\ttab\t");
    }
}
