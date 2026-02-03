//! Lazy SharedStrings parser with on-demand loading.
//!
//! This module provides a memory-efficient way to access shared strings
//! by building an index of byte offsets and parsing strings on-demand.

use std::io::{Cursor, Read};
use std::num::NonZeroUsize;
use std::sync::Arc;

use lru::LruCache;
use quick_xml::events::Event;
use quick_xml::Reader;

use crate::error::{XlexError, XlexResult};

/// Default LRU cache size for lazy shared strings.
pub const DEFAULT_CACHE_SIZE: usize = 10_000;

/// Index entry for a shared string: (byte_offset, approximate_length)
#[derive(Debug, Clone, Copy)]
struct StringIndex {
    offset: u64,
    length: u32,
}

/// Data source for shared strings.
#[derive(Debug)]
enum SharedStringsData {
    /// In-memory buffer
    InMemory(Arc<Vec<u8>>),
}

impl SharedStringsData {
    fn as_slice(&self) -> &[u8] {
        match self {
            Self::InMemory(data) => data.as_slice(),
        }
    }
}

/// Lazy parser for SharedStrings.xml with on-demand loading.
///
/// Instead of parsing all strings upfront, this parser:
/// 1. Scans the XML once to build a byte offset index
/// 2. Parses individual strings on-demand when accessed
/// 3. Caches recently accessed strings using LRU
#[derive(Debug)]
pub struct LazySharedStrings {
    /// Data source
    data: SharedStringsData,
    /// Index of string positions
    index: Vec<StringIndex>,
    /// LRU cache for recently accessed strings
    cache: LruCache<u32, String>,
    /// Total count of strings
    count: u32,
}

impl LazySharedStrings {
    /// Creates a new lazy shared strings parser from raw bytes.
    pub fn from_bytes(data: Vec<u8>, cache_size: usize) -> XlexResult<Self> {
        let data = Arc::new(data);
        let index = Self::build_index(&data)?;
        let count = index.len() as u32;

        let cache_size =
            NonZeroUsize::new(cache_size).unwrap_or(NonZeroUsize::new(DEFAULT_CACHE_SIZE).unwrap());

        Ok(Self {
            data: SharedStringsData::InMemory(data),
            index,
            cache: LruCache::new(cache_size),
            count,
        })
    }

    /// Creates a new lazy shared strings parser with default cache size.
    pub fn from_bytes_default(data: Vec<u8>) -> XlexResult<Self> {
        Self::from_bytes(data, DEFAULT_CACHE_SIZE)
    }

    /// Creates from a reader by reading all data into memory.
    pub fn from_reader<R: Read>(mut reader: R, cache_size: usize) -> XlexResult<Self> {
        let mut data = Vec::new();
        reader
            .read_to_end(&mut data)
            .map_err(|e| XlexError::IoError {
                message: e.to_string(),
                source: Some(e),
            })?;
        Self::from_bytes(data, cache_size)
    }

    /// Builds the index by scanning XML for `<si>` positions.
    fn build_index(data: &[u8]) -> XlexResult<Vec<StringIndex>> {
        let mut index = Vec::new();
        let mut reader = Reader::from_reader(Cursor::new(data));
        reader.config_mut().trim_text(false);

        let mut buf = Vec::new();
        let mut last_si_start: Option<u64> = None;

        loop {
            let pos_before = reader.buffer_position();

            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(e)) if e.name().as_ref() == b"si" => {
                    last_si_start = Some(pos_before);
                }
                Ok(Event::End(e)) if e.name().as_ref() == b"si" => {
                    if let Some(start) = last_si_start.take() {
                        let end = reader.buffer_position();
                        index.push(StringIndex {
                            offset: start,
                            length: (end - start) as u32,
                        });
                    }
                }
                Ok(Event::Eof) => break,
                Err(e) => {
                    return Err(XlexError::InvalidXml {
                        message: format!("Error building SharedStrings index: {}", e),
                    });
                }
                _ => {}
            }
            buf.clear();
        }

        Ok(index)
    }

    /// Gets a string by index, parsing on-demand if not cached.
    pub fn get(&mut self, idx: u32) -> Option<String> {
        // Check cache first
        if let Some(s) = self.cache.get(&idx) {
            return Some(s.clone());
        }

        // Parse on demand
        let string = self.parse_single(idx)?;
        self.cache.put(idx, string.clone());
        Some(string)
    }

    /// Parses a single string at the given index.
    fn parse_single(&self, idx: u32) -> Option<String> {
        let entry = self.index.get(idx as usize)?;
        let data = self.data.as_slice();

        // Get the slice containing this <si> element
        let start = entry.offset as usize;
        let end = (entry.offset + entry.length as u64) as usize;

        if end > data.len() {
            return None;
        }

        let slice = &data[start..end];
        self.parse_si_element(slice)
    }

    /// Parses a single `<si>` element to extract its text content.
    fn parse_si_element(&self, data: &[u8]) -> Option<String> {
        let mut reader = Reader::from_reader(Cursor::new(data));
        reader.config_mut().trim_text(false);

        let mut buf = Vec::new();
        let mut result = String::new();
        let mut in_si = false;
        let mut in_t = false;

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(e)) => match e.name().as_ref() {
                    b"si" => in_si = true,
                    b"t" if in_si => in_t = true,
                    _ => {}
                },
                Ok(Event::End(e)) => match e.name().as_ref() {
                    b"si" => break,
                    b"t" => in_t = false,
                    _ => {}
                },
                Ok(Event::Text(e)) => {
                    if in_t {
                        if let Ok(text) = e.unescape() {
                            result.push_str(&text);
                        }
                    }
                }
                Ok(Event::Eof) => break,
                Err(_) => break,
                _ => {}
            }
            buf.clear();
        }

        Some(result)
    }

    /// Returns the total count of strings.
    pub fn count(&self) -> u32 {
        self.count
    }

    /// Returns the number of cached strings.
    pub fn cached_count(&self) -> usize {
        self.cache.len()
    }

    /// Preloads all strings into the cache (eager mode).
    /// Use this when you know you'll access most strings.
    pub fn preload_all(&mut self) {
        for idx in 0..self.count {
            if !self.cache.contains(&idx) {
                if let Some(s) = self.parse_single(idx) {
                    self.cache.put(idx, s);
                }
            }
        }
    }

    /// Converts to a Vec of all strings (for backward compatibility).
    pub fn to_vec(&mut self) -> Vec<String> {
        (0..self.count)
            .map(|idx| self.get(idx).unwrap_or_default())
            .collect()
    }
}

/// Creates an empty LazySharedStrings for workbooks without shared strings.
impl Default for LazySharedStrings {
    fn default() -> Self {
        Self {
            data: SharedStringsData::InMemory(Arc::new(Vec::new())),
            index: Vec::new(),
            cache: LruCache::new(NonZeroUsize::new(1).unwrap()),
            count: 0,
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    const SAMPLE_SST: &[u8] = br#"<?xml version="1.0" encoding="UTF-8"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="3" uniqueCount="3">
    <si><t>Hello</t></si>
    <si><t>World</t></si>
    <si><t>Test</t></si>
</sst>"#;

    #[test]
    fn test_lazy_shared_strings_basic() {
        let mut lazy = LazySharedStrings::from_bytes_default(SAMPLE_SST.to_vec()).unwrap();

        assert_eq!(lazy.count(), 3);
        assert_eq!(lazy.get(0), Some("Hello".to_string()));
        assert_eq!(lazy.get(1), Some("World".to_string()));
        assert_eq!(lazy.get(2), Some("Test".to_string()));
        assert_eq!(lazy.get(3), None);
    }

    #[test]
    fn test_lazy_shared_strings_caching() {
        let mut lazy = LazySharedStrings::from_bytes_default(SAMPLE_SST.to_vec()).unwrap();

        assert_eq!(lazy.cached_count(), 0);
        lazy.get(0);
        assert_eq!(lazy.cached_count(), 1);
        lazy.get(0); // Should hit cache
        assert_eq!(lazy.cached_count(), 1);
        lazy.get(1);
        assert_eq!(lazy.cached_count(), 2);
    }

    #[test]
    fn test_lazy_shared_strings_to_vec() {
        let mut lazy = LazySharedStrings::from_bytes_default(SAMPLE_SST.to_vec()).unwrap();
        let vec = lazy.to_vec();

        assert_eq!(vec, vec!["Hello", "World", "Test"]);
    }

    #[test]
    fn test_lazy_shared_strings_preload() {
        let mut lazy = LazySharedStrings::from_bytes_default(SAMPLE_SST.to_vec()).unwrap();

        assert_eq!(lazy.cached_count(), 0);
        lazy.preload_all();
        assert_eq!(lazy.cached_count(), 3);
    }

    #[test]
    fn test_lazy_shared_strings_rich_text() {
        let rich_text = br#"<?xml version="1.0" encoding="UTF-8"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
    <si>
        <r><t>Bold </t></r>
        <r><t>Normal</t></r>
    </si>
</sst>"#;

        let mut lazy = LazySharedStrings::from_bytes_default(rich_text.to_vec()).unwrap();
        assert_eq!(lazy.get(0), Some("Bold Normal".to_string()));
    }

    #[test]
    fn test_lazy_shared_strings_unicode() {
        // Use regular string (not raw byte string) for Unicode content
        let unicode = r#"<?xml version="1.0" encoding="UTF-8"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
    <si><t>日本語</t></si>
    <si><t>中文</t></si>
    <si><t>🎉</t></si>
</sst>"#;

        let mut lazy = LazySharedStrings::from_bytes_default(unicode.as_bytes().to_vec()).unwrap();
        assert_eq!(lazy.get(0), Some("日本語".to_string()));
        assert_eq!(lazy.get(1), Some("中文".to_string()));
        assert_eq!(lazy.get(2), Some("🎉".to_string()));
    }

    #[test]
    fn test_lazy_shared_strings_xml_entities() {
        let entities = br#"<?xml version="1.0" encoding="UTF-8"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
    <si><t>&lt;tag&gt;</t></si>
    <si><t>A &amp; B</t></si>
</sst>"#;

        let mut lazy = LazySharedStrings::from_bytes_default(entities.to_vec()).unwrap();
        assert_eq!(lazy.get(0), Some("<tag>".to_string()));
        assert_eq!(lazy.get(1), Some("A & B".to_string()));
    }

    #[test]
    fn test_lazy_shared_strings_empty() {
        let empty = br#"<?xml version="1.0" encoding="UTF-8"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
</sst>"#;

        let mut lazy = LazySharedStrings::from_bytes_default(empty.to_vec()).unwrap();
        assert_eq!(lazy.count(), 0);
        assert_eq!(lazy.get(0), None);
    }

    #[test]
    fn test_lazy_shared_strings_default() {
        let lazy = LazySharedStrings::default();
        assert_eq!(lazy.count(), 0);
    }

    #[test]
    fn test_small_cache_eviction() {
        let mut lazy = LazySharedStrings::from_bytes(SAMPLE_SST.to_vec(), 2).unwrap();

        lazy.get(0);
        lazy.get(1);
        assert_eq!(lazy.cached_count(), 2);

        lazy.get(2); // Should evict one
        assert_eq!(lazy.cached_count(), 2);

        // All should still be accessible
        assert_eq!(lazy.get(0), Some("Hello".to_string()));
        assert_eq!(lazy.get(1), Some("World".to_string()));
        assert_eq!(lazy.get(2), Some("Test".to_string()));
    }
}
