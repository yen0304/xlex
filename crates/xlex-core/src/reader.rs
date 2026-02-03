//! Memory-efficient workbook reader with automatic mmap support.
//!
//! This module provides a reader that automatically uses memory mapping
//! for large files and falls back to buffered I/O for small files or stdin.

use std::fs::File;
use std::io::{Cursor, Read};
use std::path::{Path, PathBuf};

use memmap2::Mmap;

use crate::error::{XlexError, XlexResult};

/// Default threshold for using memory mapping (10 MB).
pub const DEFAULT_MMAP_THRESHOLD: u64 = 10 * 1024 * 1024;

/// Reader source for workbook data.
#[derive(Debug)]
enum ReaderSource {
    /// Memory-mapped file for large files
    Mmap {
        #[allow(dead_code)]
        mmap: Mmap,
        #[allow(dead_code)]
        file: File,
    },
    /// In-memory buffer for small files or stdin
    Buffer(Vec<u8>),
}

/// A memory-efficient reader for xlsx files.
///
/// Automatically selects between memory mapping and buffered I/O
/// based on file size for optimal performance.
#[derive(Debug)]
pub struct WorkbookReader {
    source: ReaderSource,
    path: Option<PathBuf>,
}

impl WorkbookReader {
    /// Opens a file with automatic mmap selection based on size.
    pub fn open(path: impl AsRef<Path>) -> XlexResult<Self> {
        Self::open_with_threshold(path, DEFAULT_MMAP_THRESHOLD)
    }

    /// Opens a file with a custom mmap threshold.
    pub fn open_with_threshold(path: impl AsRef<Path>, threshold: u64) -> XlexResult<Self> {
        let path = path.as_ref();

        // Check file exists
        if !path.exists() {
            return Err(XlexError::FileNotFound {
                path: path.to_path_buf(),
            });
        }

        let metadata = std::fs::metadata(path).map_err(|e| XlexError::IoError {
            message: format!("Failed to read file metadata: {}", e),
            source: Some(e),
        })?;

        let file = File::open(path).map_err(|e| {
            if e.kind() == std::io::ErrorKind::PermissionDenied {
                XlexError::PermissionDenied {
                    path: path.to_path_buf(),
                }
            } else {
                XlexError::IoError {
                    message: e.to_string(),
                    source: Some(e),
                }
            }
        })?;

        let source = if metadata.len() > threshold {
            // Use memory mapping for large files
            // SAFETY: We keep the file handle open for the lifetime of the mmap
            let mmap = unsafe { Mmap::map(&file) }.map_err(|e| XlexError::IoError {
                message: format!("Failed to memory-map file: {}", e),
                source: None,
            })?;
            ReaderSource::Mmap { mmap, file }
        } else {
            // Read into memory for small files
            let mut file = file;
            let mut data = Vec::with_capacity(metadata.len() as usize);
            file.read_to_end(&mut data)
                .map_err(|e| XlexError::IoError {
                    message: e.to_string(),
                    source: Some(e),
                })?;
            ReaderSource::Buffer(data)
        };

        Ok(Self {
            source,
            path: Some(path.to_path_buf()),
        })
    }

    /// Creates a reader from bytes (for stdin or testing).
    pub fn from_bytes(data: Vec<u8>) -> Self {
        Self {
            source: ReaderSource::Buffer(data),
            path: None,
        }
    }

    /// Creates a reader from any Read source.
    pub fn from_reader<R: Read>(mut reader: R) -> XlexResult<Self> {
        let mut data = Vec::new();
        reader
            .read_to_end(&mut data)
            .map_err(|e| XlexError::IoError {
                message: e.to_string(),
                source: Some(e),
            })?;
        Ok(Self::from_bytes(data))
    }

    /// Returns the file path if available.
    pub fn path(&self) -> Option<&Path> {
        self.path.as_deref()
    }

    /// Returns the data as a byte slice.
    pub fn as_bytes(&self) -> &[u8] {
        match &self.source {
            ReaderSource::Mmap { mmap, .. } => mmap.as_ref(),
            ReaderSource::Buffer(data) => data.as_slice(),
        }
    }

    /// Returns the length of the data.
    pub fn len(&self) -> usize {
        self.as_bytes().len()
    }

    /// Returns true if the reader is empty.
    pub fn is_empty(&self) -> bool {
        self.len() == 0
    }

    /// Returns true if using memory mapping.
    pub fn is_mmap(&self) -> bool {
        matches!(self.source, ReaderSource::Mmap { .. })
    }

    /// Creates a cursor for reading the data.
    pub fn cursor(&self) -> Cursor<&[u8]> {
        Cursor::new(self.as_bytes())
    }
}

impl AsRef<[u8]> for WorkbookReader {
    fn as_ref(&self) -> &[u8] {
        self.as_bytes()
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::io::Write;
    use tempfile::NamedTempFile;

    #[test]
    fn test_reader_from_bytes() {
        let data = vec![1, 2, 3, 4, 5];
        let reader = WorkbookReader::from_bytes(data.clone());

        assert_eq!(reader.as_bytes(), &data);
        assert_eq!(reader.len(), 5);
        assert!(!reader.is_empty());
        assert!(!reader.is_mmap());
        assert!(reader.path().is_none());
    }

    #[test]
    fn test_reader_small_file() {
        let mut temp = NamedTempFile::new().unwrap();
        temp.write_all(b"small file content").unwrap();
        temp.flush().unwrap();

        let reader = WorkbookReader::open(temp.path()).unwrap();

        assert!(!reader.is_mmap()); // Small file should use buffer
        assert_eq!(reader.as_bytes(), b"small file content");
        assert!(reader.path().is_some());
    }

    #[test]
    fn test_reader_threshold() {
        let mut temp = NamedTempFile::new().unwrap();
        temp.write_all(b"small").unwrap();
        temp.flush().unwrap();

        // With threshold of 1 byte, even small files use mmap
        let reader = WorkbookReader::open_with_threshold(temp.path(), 1).unwrap();
        assert!(reader.is_mmap());

        // With threshold of 1MB, small file uses buffer
        let reader = WorkbookReader::open_with_threshold(temp.path(), 1024 * 1024).unwrap();
        assert!(!reader.is_mmap());
    }

    #[test]
    fn test_reader_nonexistent_file() {
        let result = WorkbookReader::open("/nonexistent/path/file.xlsx");
        assert!(result.is_err());
    }

    #[test]
    fn test_reader_cursor() {
        let data = vec![1, 2, 3, 4, 5];
        let reader = WorkbookReader::from_bytes(data);

        let mut cursor = reader.cursor();
        let mut buf = [0u8; 3];
        cursor.read_exact(&mut buf).unwrap();
        assert_eq!(buf, [1, 2, 3]);
    }

    #[test]
    fn test_reader_as_ref() {
        let data = vec![1, 2, 3];
        let reader = WorkbookReader::from_bytes(data.clone());

        let slice: &[u8] = reader.as_ref();
        assert_eq!(slice, &data);
    }
}
