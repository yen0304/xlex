//! Error types for xlex-core.
//!
//! All errors include machine-readable codes (XLEX_E001-E099) and human-friendly messages.

use std::path::PathBuf;
use thiserror::Error;

/// Result type alias for xlex operations.
pub type XlexResult<T> = Result<T, XlexError>;

/// Error codes for XLEX operations.
///
/// Format: XLEX_EXXX where XXX is a three-digit number.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ErrorCode {
    // File errors (E001-E009)
    FileNotFound = 1,
    FileExists = 2,
    PermissionDenied = 3,
    InvalidExtension = 4,
    IoError = 5,

    // Parse errors (E010-E019)
    ParseError = 10,
    InvalidZipStructure = 11,
    MissingRequiredEntry = 12,
    InvalidXml = 13,
    EncodingError = 14,

    // Reference errors (E020-E029)
    InvalidReference = 20,
    InvalidRange = 21,
    ReferenceOutOfBounds = 22,

    // Sheet errors (E030-E039)
    SheetNotFound = 30,
    SheetAlreadyExists = 31,
    InvalidSheetName = 32,
    SheetIndexOutOfBounds = 33,
    CannotDeleteLastSheet = 34,

    // Cell errors (E040-E049)
    CellNotFound = 40,
    InvalidCellValue = 41,
    InvalidFormula = 42,
    CircularReference = 43,

    // Style errors (E050-E059)
    StyleNotFound = 50,
    InvalidStyle = 51,

    // Operation errors (E060-E069)
    OperationFailed = 60,
    InvalidOperation = 61,
    UnsupportedOperation = 62,

    // Template errors (E070-E079)
    TemplateParseError = 70,
    TemplateRenderError = 71,
    InvalidTemplateData = 72,

    // Config errors (E080-E089)
    ConfigError = 80,
    InvalidConfig = 81,

    // General errors (E090-E099)
    InternalError = 90,
    NotImplemented = 99,
}

impl ErrorCode {
    /// Returns the error code as a string (e.g., "XLEX_E001").
    pub fn as_str(&self) -> String {
        format!("XLEX_E{:03}", *self as u32)
    }
}

impl std::fmt::Display for ErrorCode {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        write!(f, "{}", self.as_str())
    }
}

/// Main error type for xlex operations.
#[derive(Error, Debug)]
pub enum XlexError {
    // File errors
    #[error("{}: File not found: {path:?}", ErrorCode::FileNotFound)]
    FileNotFound { path: PathBuf },

    #[error("{}: File already exists: {path:?}", ErrorCode::FileExists)]
    FileExists { path: PathBuf },

    #[error("{}: Permission denied: {path:?}", ErrorCode::PermissionDenied)]
    PermissionDenied { path: PathBuf },

    #[error("{}: Invalid file extension, expected .xlsx: {path:?}", ErrorCode::InvalidExtension)]
    InvalidExtension { path: PathBuf },

    #[error("{}: I/O error: {message}", ErrorCode::IoError)]
    IoError {
        message: String,
        #[source]
        source: Option<std::io::Error>,
    },

    // Parse errors
    #[error("{}: Parse error at {location}: {message}", ErrorCode::ParseError)]
    ParseError { message: String, location: String },

    #[error("{}: Invalid ZIP structure: {message}", ErrorCode::InvalidZipStructure)]
    InvalidZipStructure { message: String },

    #[error("{}: Missing required entry: {entry}", ErrorCode::MissingRequiredEntry)]
    MissingRequiredEntry { entry: String },

    #[error("{}: Invalid XML: {message}", ErrorCode::InvalidXml)]
    InvalidXml { message: String },

    #[error("{}: Encoding error: {message}", ErrorCode::EncodingError)]
    EncodingError { message: String },

    // Reference errors
    #[error("{}: Invalid cell reference: {reference}", ErrorCode::InvalidReference)]
    InvalidReference { reference: String },

    #[error("{}: Invalid range: {range}", ErrorCode::InvalidRange)]
    InvalidRange { range: String },

    #[error("{}: Reference out of bounds: {reference}", ErrorCode::ReferenceOutOfBounds)]
    ReferenceOutOfBounds { reference: String },

    // Sheet errors
    #[error("{}: Sheet not found: {name}", ErrorCode::SheetNotFound)]
    SheetNotFound { name: String },

    #[error("{}: Sheet already exists: {name}", ErrorCode::SheetAlreadyExists)]
    SheetAlreadyExists { name: String },

    #[error("{}: Invalid sheet name: {name} ({reason})", ErrorCode::InvalidSheetName)]
    InvalidSheetName { name: String, reason: String },

    #[error("{}: Sheet index out of bounds: {index}", ErrorCode::SheetIndexOutOfBounds)]
    SheetIndexOutOfBounds { index: usize },

    #[error("{}: Cannot delete the last sheet", ErrorCode::CannotDeleteLastSheet)]
    CannotDeleteLastSheet,

    // Cell errors
    #[error("{}: Cell not found: {reference}", ErrorCode::CellNotFound)]
    CellNotFound { reference: String },

    #[error("{}: Invalid cell value: {message}", ErrorCode::InvalidCellValue)]
    InvalidCellValue { message: String },

    #[error("{}: Invalid formula: {formula} ({reason})", ErrorCode::InvalidFormula)]
    InvalidFormula { formula: String, reason: String },

    #[error("{}: Circular reference detected: {path}", ErrorCode::CircularReference)]
    CircularReference { path: String },

    // Style errors
    #[error("{}: Style not found: {id}", ErrorCode::StyleNotFound)]
    StyleNotFound { id: u32 },

    #[error("{}: Invalid style: {message}", ErrorCode::InvalidStyle)]
    InvalidStyle { message: String },

    // Operation errors
    #[error("{}: Operation failed: {message}", ErrorCode::OperationFailed)]
    OperationFailed { message: String },

    #[error("{}: Invalid operation: {message}", ErrorCode::InvalidOperation)]
    InvalidOperation { message: String },

    #[error("{}: Unsupported operation: {message}", ErrorCode::UnsupportedOperation)]
    UnsupportedOperation { message: String },

    // Template errors
    #[error("{}: Template parse error: {message}", ErrorCode::TemplateParseError)]
    TemplateParseError { message: String },

    #[error("{}: Template render error: {message}", ErrorCode::TemplateRenderError)]
    TemplateRenderError { message: String },

    #[error("{}: Invalid template data: {message}", ErrorCode::InvalidTemplateData)]
    InvalidTemplateData { message: String },

    // Config errors
    #[error("{}: Configuration error: {message}", ErrorCode::ConfigError)]
    ConfigError { message: String },

    #[error("{}: Invalid configuration: {message}", ErrorCode::InvalidConfig)]
    InvalidConfig { message: String },

    // General errors
    #[error("{}: Internal error: {message}", ErrorCode::InternalError)]
    InternalError { message: String },

    #[error("{}: Not implemented: {feature}", ErrorCode::NotImplemented)]
    NotImplemented { feature: String },
}

impl XlexError {
    /// Returns the error code for this error.
    pub fn code(&self) -> ErrorCode {
        match self {
            XlexError::FileNotFound { .. } => ErrorCode::FileNotFound,
            XlexError::FileExists { .. } => ErrorCode::FileExists,
            XlexError::PermissionDenied { .. } => ErrorCode::PermissionDenied,
            XlexError::InvalidExtension { .. } => ErrorCode::InvalidExtension,
            XlexError::IoError { .. } => ErrorCode::IoError,
            XlexError::ParseError { .. } => ErrorCode::ParseError,
            XlexError::InvalidZipStructure { .. } => ErrorCode::InvalidZipStructure,
            XlexError::MissingRequiredEntry { .. } => ErrorCode::MissingRequiredEntry,
            XlexError::InvalidXml { .. } => ErrorCode::InvalidXml,
            XlexError::EncodingError { .. } => ErrorCode::EncodingError,
            XlexError::InvalidReference { .. } => ErrorCode::InvalidReference,
            XlexError::InvalidRange { .. } => ErrorCode::InvalidRange,
            XlexError::ReferenceOutOfBounds { .. } => ErrorCode::ReferenceOutOfBounds,
            XlexError::SheetNotFound { .. } => ErrorCode::SheetNotFound,
            XlexError::SheetAlreadyExists { .. } => ErrorCode::SheetAlreadyExists,
            XlexError::InvalidSheetName { .. } => ErrorCode::InvalidSheetName,
            XlexError::SheetIndexOutOfBounds { .. } => ErrorCode::SheetIndexOutOfBounds,
            XlexError::CannotDeleteLastSheet => ErrorCode::CannotDeleteLastSheet,
            XlexError::CellNotFound { .. } => ErrorCode::CellNotFound,
            XlexError::InvalidCellValue { .. } => ErrorCode::InvalidCellValue,
            XlexError::InvalidFormula { .. } => ErrorCode::InvalidFormula,
            XlexError::CircularReference { .. } => ErrorCode::CircularReference,
            XlexError::StyleNotFound { .. } => ErrorCode::StyleNotFound,
            XlexError::InvalidStyle { .. } => ErrorCode::InvalidStyle,
            XlexError::OperationFailed { .. } => ErrorCode::OperationFailed,
            XlexError::InvalidOperation { .. } => ErrorCode::InvalidOperation,
            XlexError::UnsupportedOperation { .. } => ErrorCode::UnsupportedOperation,
            XlexError::TemplateParseError { .. } => ErrorCode::TemplateParseError,
            XlexError::TemplateRenderError { .. } => ErrorCode::TemplateRenderError,
            XlexError::InvalidTemplateData { .. } => ErrorCode::InvalidTemplateData,
            XlexError::ConfigError { .. } => ErrorCode::ConfigError,
            XlexError::InvalidConfig { .. } => ErrorCode::InvalidConfig,
            XlexError::InternalError { .. } => ErrorCode::InternalError,
            XlexError::NotImplemented { .. } => ErrorCode::NotImplemented,
        }
    }

    /// Returns the exit code to use when this error causes program termination.
    pub fn exit_code(&self) -> i32 {
        self.code() as i32
    }
}

impl From<std::io::Error> for XlexError {
    fn from(err: std::io::Error) -> Self {
        XlexError::IoError {
            message: err.to_string(),
            source: Some(err),
        }
    }
}

impl From<zip::result::ZipError> for XlexError {
    fn from(err: zip::result::ZipError) -> Self {
        XlexError::InvalidZipStructure {
            message: err.to_string(),
        }
    }
}

impl From<quick_xml::Error> for XlexError {
    fn from(err: quick_xml::Error) -> Self {
        XlexError::InvalidXml {
            message: err.to_string(),
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_error_code_display() {
        assert_eq!(ErrorCode::FileNotFound.as_str(), "XLEX_E001");
        assert_eq!(ErrorCode::ParseError.as_str(), "XLEX_E010");
        assert_eq!(ErrorCode::InvalidReference.as_str(), "XLEX_E020");
    }

    #[test]
    fn test_error_code_from_error() {
        let err = XlexError::FileNotFound {
            path: PathBuf::from("test.xlsx"),
        };
        assert_eq!(err.code(), ErrorCode::FileNotFound);
        assert_eq!(err.exit_code(), 1);
    }
}
