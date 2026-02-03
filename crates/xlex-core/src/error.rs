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

    #[error(
        "{}: Invalid file extension, expected .xlsx: {path:?}",
        ErrorCode::InvalidExtension
    )]
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

    #[error(
        "{}: Reference out of bounds: {reference}",
        ErrorCode::ReferenceOutOfBounds
    )]
    ReferenceOutOfBounds { reference: String },

    // Sheet errors
    #[error("{}: Sheet not found: {name}", ErrorCode::SheetNotFound)]
    SheetNotFound { name: String },

    #[error("{}: Sheet already exists: {name}", ErrorCode::SheetAlreadyExists)]
    SheetAlreadyExists { name: String },

    #[error(
        "{}: Invalid sheet name: {name} ({reason})",
        ErrorCode::InvalidSheetName
    )]
    InvalidSheetName { name: String, reason: String },

    #[error(
        "{}: Sheet index out of bounds: {index}",
        ErrorCode::SheetIndexOutOfBounds
    )]
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

    #[error(
        "{}: Circular reference detected: {path}",
        ErrorCode::CircularReference
    )]
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

    #[error(
        "{}: Unsupported operation: {message}",
        ErrorCode::UnsupportedOperation
    )]
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

    /// Returns a recovery suggestion for this error.
    ///
    /// Provides actionable advice to help users resolve common issues.
    pub fn recovery_suggestion(&self) -> Option<&'static str> {
        match self {
            XlexError::FileNotFound { .. } => Some(
                "Check if the file path is correct. Use `ls` to verify the file exists.",
            ),
            XlexError::FileExists { .. } => Some(
                "Use --force to overwrite the existing file, or choose a different output path.",
            ),
            XlexError::PermissionDenied { .. } => Some(
                "Check file permissions with `ls -la`. You may need to use `chmod` or run with elevated privileges.",
            ),
            XlexError::InvalidExtension { .. } => Some(
                "Ensure the file has a .xlsx extension. Use `xlex convert` to convert from other formats.",
            ),
            XlexError::IoError { .. } => Some(
                "Check disk space and file system permissions. The file may be in use by another process.",
            ),
            XlexError::ParseError { .. } => Some(
                "The file may be corrupted. Try opening it in Excel to verify, or use a backup.",
            ),
            XlexError::InvalidZipStructure { .. } => Some(
                "The xlsx file appears to be corrupted. Try re-downloading or restoring from backup.",
            ),
            XlexError::MissingRequiredEntry { .. } => Some(
                "The xlsx file is missing required components. It may be corrupted or not a valid xlsx file.",
            ),
            XlexError::InvalidXml { .. } => Some(
                "The file contains invalid XML. It may have been modified by a non-Excel application.",
            ),
            XlexError::InvalidReference { .. } => Some(
                "Use A1 notation (e.g., A1, B2, AA100). Column letters are A-XFD, rows are 1-1048576.",
            ),
            XlexError::InvalidRange { .. } => Some(
                "Use range notation like A1:B10. Start cell should be top-left, end cell bottom-right.",
            ),
            XlexError::ReferenceOutOfBounds { .. } => Some(
                "Excel sheets support columns A-XFD (16384) and rows 1-1048576.",
            ),
            XlexError::SheetNotFound { .. } => Some(
                "Use `xlex sheet list <file>` to see available sheet names.",
            ),
            XlexError::SheetAlreadyExists { .. } => Some(
                "Choose a different sheet name, or use `xlex sheet remove` to delete the existing one.",
            ),
            XlexError::InvalidSheetName { .. } => Some(
                "Sheet names cannot contain: \\ / ? * [ ] : and cannot exceed 31 characters.",
            ),
            XlexError::SheetIndexOutOfBounds { .. } => Some(
                "Use `xlex sheet list <file>` to see the number of available sheets.",
            ),
            XlexError::CannotDeleteLastSheet => Some(
                "A workbook must have at least one sheet. Add a new sheet before deleting this one.",
            ),
            XlexError::CellNotFound { .. } => Some(
                "The cell may be empty. Use `xlex cell get` to check the cell value.",
            ),
            XlexError::InvalidCellValue { .. } => Some(
                "Check the value format. Numbers, text, booleans, and formulas (starting with =) are supported.",
            ),
            XlexError::InvalidFormula { .. } => Some(
                "Formulas must start with '='. Use `xlex formula validate` to check formula syntax.",
            ),
            XlexError::CircularReference { .. } => Some(
                "Remove the circular dependency. Use `xlex formula refs` to trace cell dependencies.",
            ),
            XlexError::StyleNotFound { .. } => Some(
                "Use `xlex style list <file>` to see available style IDs.",
            ),
            XlexError::InvalidStyle { .. } => Some(
                "Check the style specification. Use `xlex style preset list` for preset styles.",
            ),
            XlexError::TemplateParseError { .. } => Some(
                "Check template syntax. Use {{variable}}, {{#each items}}, {{#if condition}}.",
            ),
            XlexError::TemplateRenderError { .. } => Some(
                "Ensure all template variables are provided in the data. Use `xlex template validate`.",
            ),
            XlexError::InvalidTemplateData { .. } => Some(
                "Data must be valid JSON or YAML. Use `jq .` or `yq .` to validate.",
            ),
            XlexError::ConfigError { .. } => Some(
                "Check your .xlex.yml configuration. Use `xlex config validate` to check for errors.",
            ),
            XlexError::InvalidConfig { .. } => Some(
                "Review the configuration file syntax. Use `xlex config init` to create a valid template.",
            ),
            XlexError::OperationFailed { .. } | XlexError::InvalidOperation { .. } => Some(
                "Check the command syntax with `xlex <command> --help` for usage information.",
            ),
            XlexError::UnsupportedOperation { .. } => Some(
                "This operation is not yet supported. Check the documentation for alternatives.",
            ),
            XlexError::InternalError { .. } => Some(
                "Please report this issue at https://github.com/xlex/xlex/issues with the full error message.",
            ),
            XlexError::NotImplemented { .. } => Some(
                "This feature is not yet implemented. Check upcoming releases or contribute!",
            ),
            _ => None,
        }
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
        assert_eq!(ErrorCode::SheetNotFound.as_str(), "XLEX_E030");
        assert_eq!(ErrorCode::CellNotFound.as_str(), "XLEX_E040");
        assert_eq!(ErrorCode::StyleNotFound.as_str(), "XLEX_E050");
        assert_eq!(ErrorCode::OperationFailed.as_str(), "XLEX_E060");
        assert_eq!(ErrorCode::TemplateParseError.as_str(), "XLEX_E070");
        assert_eq!(ErrorCode::ConfigError.as_str(), "XLEX_E080");
        assert_eq!(ErrorCode::InternalError.as_str(), "XLEX_E090");
        assert_eq!(ErrorCode::NotImplemented.as_str(), "XLEX_E099");
    }

    #[test]
    fn test_error_code_display_trait() {
        assert_eq!(format!("{}", ErrorCode::FileNotFound), "XLEX_E001");
        assert_eq!(format!("{}", ErrorCode::IoError), "XLEX_E005");
    }

    #[test]
    fn test_error_code_from_error() {
        let err = XlexError::FileNotFound {
            path: PathBuf::from("test.xlsx"),
        };
        assert_eq!(err.code(), ErrorCode::FileNotFound);
        assert_eq!(err.exit_code(), 1);
    }

    #[test]
    fn test_all_error_codes() {
        // File errors
        assert_eq!(
            XlexError::FileNotFound {
                path: PathBuf::from("test")
            }
            .code(),
            ErrorCode::FileNotFound
        );
        assert_eq!(
            XlexError::FileExists {
                path: PathBuf::from("test")
            }
            .code(),
            ErrorCode::FileExists
        );
        assert_eq!(
            XlexError::PermissionDenied {
                path: PathBuf::from("test")
            }
            .code(),
            ErrorCode::PermissionDenied
        );
        assert_eq!(
            XlexError::InvalidExtension {
                path: PathBuf::from("test")
            }
            .code(),
            ErrorCode::InvalidExtension
        );
        assert_eq!(
            XlexError::IoError {
                message: "test".to_string(),
                source: None
            }
            .code(),
            ErrorCode::IoError
        );

        // Parse errors
        assert_eq!(
            XlexError::ParseError {
                message: "test".to_string(),
                location: "test".to_string()
            }
            .code(),
            ErrorCode::ParseError
        );
        assert_eq!(
            XlexError::InvalidZipStructure {
                message: "test".to_string()
            }
            .code(),
            ErrorCode::InvalidZipStructure
        );
        assert_eq!(
            XlexError::MissingRequiredEntry {
                entry: "test".to_string()
            }
            .code(),
            ErrorCode::MissingRequiredEntry
        );
        assert_eq!(
            XlexError::InvalidXml {
                message: "test".to_string()
            }
            .code(),
            ErrorCode::InvalidXml
        );
        assert_eq!(
            XlexError::EncodingError {
                message: "test".to_string()
            }
            .code(),
            ErrorCode::EncodingError
        );

        // Reference errors
        assert_eq!(
            XlexError::InvalidReference {
                reference: "test".to_string()
            }
            .code(),
            ErrorCode::InvalidReference
        );
        assert_eq!(
            XlexError::InvalidRange {
                range: "test".to_string()
            }
            .code(),
            ErrorCode::InvalidRange
        );
        assert_eq!(
            XlexError::ReferenceOutOfBounds {
                reference: "test".to_string()
            }
            .code(),
            ErrorCode::ReferenceOutOfBounds
        );

        // Sheet errors
        assert_eq!(
            XlexError::SheetNotFound {
                name: "test".to_string()
            }
            .code(),
            ErrorCode::SheetNotFound
        );
        assert_eq!(
            XlexError::SheetAlreadyExists {
                name: "test".to_string()
            }
            .code(),
            ErrorCode::SheetAlreadyExists
        );
        assert_eq!(
            XlexError::InvalidSheetName {
                name: "test".to_string(),
                reason: "test".to_string()
            }
            .code(),
            ErrorCode::InvalidSheetName
        );
        assert_eq!(
            XlexError::SheetIndexOutOfBounds { index: 0 }.code(),
            ErrorCode::SheetIndexOutOfBounds
        );
        assert_eq!(
            XlexError::CannotDeleteLastSheet.code(),
            ErrorCode::CannotDeleteLastSheet
        );

        // Cell errors
        assert_eq!(
            XlexError::CellNotFound {
                reference: "A1".to_string()
            }
            .code(),
            ErrorCode::CellNotFound
        );
        assert_eq!(
            XlexError::InvalidCellValue {
                message: "test".to_string()
            }
            .code(),
            ErrorCode::InvalidCellValue
        );
        assert_eq!(
            XlexError::InvalidFormula {
                formula: "test".to_string(),
                reason: "test".to_string()
            }
            .code(),
            ErrorCode::InvalidFormula
        );
        assert_eq!(
            XlexError::CircularReference {
                path: "A1->B1->A1".to_string()
            }
            .code(),
            ErrorCode::CircularReference
        );

        // Style errors
        assert_eq!(
            XlexError::StyleNotFound { id: 0 }.code(),
            ErrorCode::StyleNotFound
        );
        assert_eq!(
            XlexError::InvalidStyle {
                message: "test".to_string()
            }
            .code(),
            ErrorCode::InvalidStyle
        );

        // Operation errors
        assert_eq!(
            XlexError::OperationFailed {
                message: "test".to_string()
            }
            .code(),
            ErrorCode::OperationFailed
        );
        assert_eq!(
            XlexError::InvalidOperation {
                message: "test".to_string()
            }
            .code(),
            ErrorCode::InvalidOperation
        );
        assert_eq!(
            XlexError::UnsupportedOperation {
                message: "test".to_string()
            }
            .code(),
            ErrorCode::UnsupportedOperation
        );

        // Template errors
        assert_eq!(
            XlexError::TemplateParseError {
                message: "test".to_string()
            }
            .code(),
            ErrorCode::TemplateParseError
        );
        assert_eq!(
            XlexError::TemplateRenderError {
                message: "test".to_string()
            }
            .code(),
            ErrorCode::TemplateRenderError
        );
        assert_eq!(
            XlexError::InvalidTemplateData {
                message: "test".to_string()
            }
            .code(),
            ErrorCode::InvalidTemplateData
        );

        // Config errors
        assert_eq!(
            XlexError::ConfigError {
                message: "test".to_string()
            }
            .code(),
            ErrorCode::ConfigError
        );
        assert_eq!(
            XlexError::InvalidConfig {
                message: "test".to_string()
            }
            .code(),
            ErrorCode::InvalidConfig
        );

        // General errors
        assert_eq!(
            XlexError::InternalError {
                message: "test".to_string()
            }
            .code(),
            ErrorCode::InternalError
        );
        assert_eq!(
            XlexError::NotImplemented {
                feature: "test".to_string()
            }
            .code(),
            ErrorCode::NotImplemented
        );
    }

    #[test]
    fn test_recovery_suggestions() {
        // Test that all errors have proper suggestions
        assert!(XlexError::FileNotFound {
            path: PathBuf::from("test")
        }
        .recovery_suggestion()
        .is_some());
        assert!(XlexError::FileExists {
            path: PathBuf::from("test")
        }
        .recovery_suggestion()
        .is_some());
        assert!(XlexError::PermissionDenied {
            path: PathBuf::from("test")
        }
        .recovery_suggestion()
        .is_some());
        assert!(XlexError::InvalidExtension {
            path: PathBuf::from("test")
        }
        .recovery_suggestion()
        .is_some());
        assert!(XlexError::IoError {
            message: "test".to_string(),
            source: None
        }
        .recovery_suggestion()
        .is_some());
        assert!(XlexError::ParseError {
            message: "test".to_string(),
            location: "test".to_string()
        }
        .recovery_suggestion()
        .is_some());
        assert!(XlexError::InvalidZipStructure {
            message: "test".to_string()
        }
        .recovery_suggestion()
        .is_some());
        assert!(XlexError::MissingRequiredEntry {
            entry: "test".to_string()
        }
        .recovery_suggestion()
        .is_some());
        assert!(XlexError::InvalidXml {
            message: "test".to_string()
        }
        .recovery_suggestion()
        .is_some());
        assert!(XlexError::InvalidReference {
            reference: "test".to_string()
        }
        .recovery_suggestion()
        .is_some());
        assert!(XlexError::InvalidRange {
            range: "test".to_string()
        }
        .recovery_suggestion()
        .is_some());
        assert!(XlexError::ReferenceOutOfBounds {
            reference: "test".to_string()
        }
        .recovery_suggestion()
        .is_some());
        assert!(XlexError::SheetNotFound {
            name: "test".to_string()
        }
        .recovery_suggestion()
        .is_some());
        assert!(XlexError::SheetAlreadyExists {
            name: "test".to_string()
        }
        .recovery_suggestion()
        .is_some());
        assert!(XlexError::InvalidSheetName {
            name: "test".to_string(),
            reason: "test".to_string()
        }
        .recovery_suggestion()
        .is_some());
        assert!(XlexError::SheetIndexOutOfBounds { index: 0 }
            .recovery_suggestion()
            .is_some());
        assert!(XlexError::CannotDeleteLastSheet
            .recovery_suggestion()
            .is_some());
        assert!(XlexError::CellNotFound {
            reference: "A1".to_string()
        }
        .recovery_suggestion()
        .is_some());
        assert!(XlexError::InvalidCellValue {
            message: "test".to_string()
        }
        .recovery_suggestion()
        .is_some());
        assert!(XlexError::InvalidFormula {
            formula: "test".to_string(),
            reason: "test".to_string()
        }
        .recovery_suggestion()
        .is_some());
        assert!(XlexError::CircularReference {
            path: "A1->B1".to_string()
        }
        .recovery_suggestion()
        .is_some());
        assert!(XlexError::StyleNotFound { id: 0 }
            .recovery_suggestion()
            .is_some());
        assert!(XlexError::InvalidStyle {
            message: "test".to_string()
        }
        .recovery_suggestion()
        .is_some());
        assert!(XlexError::TemplateParseError {
            message: "test".to_string()
        }
        .recovery_suggestion()
        .is_some());
        assert!(XlexError::TemplateRenderError {
            message: "test".to_string()
        }
        .recovery_suggestion()
        .is_some());
        assert!(XlexError::InvalidTemplateData {
            message: "test".to_string()
        }
        .recovery_suggestion()
        .is_some());
        assert!(XlexError::ConfigError {
            message: "test".to_string()
        }
        .recovery_suggestion()
        .is_some());
        assert!(XlexError::InvalidConfig {
            message: "test".to_string()
        }
        .recovery_suggestion()
        .is_some());
        assert!(XlexError::OperationFailed {
            message: "test".to_string()
        }
        .recovery_suggestion()
        .is_some());
        assert!(XlexError::InvalidOperation {
            message: "test".to_string()
        }
        .recovery_suggestion()
        .is_some());
        assert!(XlexError::UnsupportedOperation {
            message: "test".to_string()
        }
        .recovery_suggestion()
        .is_some());
        assert!(XlexError::InternalError {
            message: "test".to_string()
        }
        .recovery_suggestion()
        .is_some());
        assert!(XlexError::NotImplemented {
            feature: "test".to_string()
        }
        .recovery_suggestion()
        .is_some());
    }

    #[test]
    fn test_error_display_messages() {
        // Verify error messages contain the error code and relevant info
        let err = XlexError::FileNotFound {
            path: PathBuf::from("test.xlsx"),
        };
        let msg = format!("{}", err);
        assert!(msg.contains("XLEX_E001"));
        assert!(msg.contains("test.xlsx"));

        let err = XlexError::SheetNotFound {
            name: "MySheet".to_string(),
        };
        let msg = format!("{}", err);
        assert!(msg.contains("XLEX_E030"));
        assert!(msg.contains("MySheet"));

        let err = XlexError::InvalidFormula {
            formula: "=SUM(A1".to_string(),
            reason: "missing closing parenthesis".to_string(),
        };
        let msg = format!("{}", err);
        assert!(msg.contains("XLEX_E042"));
        assert!(msg.contains("=SUM(A1"));
        assert!(msg.contains("missing closing parenthesis"));
    }

    #[test]
    fn test_from_io_error() {
        let io_err = std::io::Error::new(std::io::ErrorKind::NotFound, "file not found");
        let xlex_err: XlexError = io_err.into();
        assert_eq!(xlex_err.code(), ErrorCode::IoError);
        let msg = format!("{}", xlex_err);
        assert!(msg.contains("file not found"));
    }

    #[test]
    fn test_exit_codes() {
        // Verify exit codes match error code values
        assert_eq!(
            XlexError::FileNotFound {
                path: PathBuf::from("test")
            }
            .exit_code(),
            1
        );
        assert_eq!(
            XlexError::ParseError {
                message: "test".to_string(),
                location: "test".to_string()
            }
            .exit_code(),
            10
        );
        assert_eq!(
            XlexError::InvalidReference {
                reference: "test".to_string()
            }
            .exit_code(),
            20
        );
        assert_eq!(
            XlexError::SheetNotFound {
                name: "test".to_string()
            }
            .exit_code(),
            30
        );
        assert_eq!(
            XlexError::CellNotFound {
                reference: "test".to_string()
            }
            .exit_code(),
            40
        );
        assert_eq!(XlexError::StyleNotFound { id: 0 }.exit_code(), 50);
        assert_eq!(
            XlexError::OperationFailed {
                message: "test".to_string()
            }
            .exit_code(),
            60
        );
        assert_eq!(
            XlexError::TemplateParseError {
                message: "test".to_string()
            }
            .exit_code(),
            70
        );
        assert_eq!(
            XlexError::ConfigError {
                message: "test".to_string()
            }
            .exit_code(),
            80
        );
        assert_eq!(
            XlexError::InternalError {
                message: "test".to_string()
            }
            .exit_code(),
            90
        );
        assert_eq!(
            XlexError::NotImplemented {
                feature: "test".to_string()
            }
            .exit_code(),
            99
        );
    }
}
