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
}
