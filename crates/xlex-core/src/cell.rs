//! Cell types and operations.

use serde::{Deserialize, Serialize};
use std::fmt;
use std::str::FromStr;

use crate::error::{XlexError, XlexResult};

/// A reference to a cell in A1 notation.
#[derive(Debug, Clone, PartialEq, Eq, Hash, Serialize, Deserialize)]
pub struct CellRef {
    /// Column number (1-indexed, A=1, B=2, ..., XFD=16384)
    pub col: u32,
    /// Row number (1-indexed, 1-1048576)
    pub row: u32,
}

impl CellRef {
    /// Maximum column number (XFD = 16384)
    pub const MAX_COL: u32 = 16384;
    /// Maximum row number
    pub const MAX_ROW: u32 = 1048576;

    /// Creates a new cell reference.
    ///
    /// # Arguments
    /// * `col` - Column number (1-indexed)
    /// * `row` - Row number (1-indexed)
    pub fn new(col: u32, row: u32) -> Self {
        Self { col, row }
    }

    /// Parses an A1-style reference (e.g., "A1", "AA100", "XFD1048576").
    pub fn parse(s: &str) -> XlexResult<Self> {
        let s = s.trim().to_uppercase();
        if s.is_empty() {
            return Err(XlexError::InvalidReference {
                reference: s.to_string(),
            });
        }

        // Find the split point between letters and digits
        let letter_end = s.chars().take_while(|c| c.is_ascii_alphabetic()).count();
        if letter_end == 0 || letter_end == s.len() {
            return Err(XlexError::InvalidReference {
                reference: s.to_string(),
            });
        }

        let col_str = &s[..letter_end];
        let row_str = &s[letter_end..];

        // Parse column
        let col = Self::col_from_letters(col_str).ok_or_else(|| XlexError::InvalidReference {
            reference: s.to_string(),
        })?;

        // Parse row
        let row: u32 = row_str.parse().map_err(|_| XlexError::InvalidReference {
            reference: s.to_string(),
        })?;

        // Validate bounds
        if col == 0 || col > Self::MAX_COL {
            return Err(XlexError::InvalidReference {
                reference: s.to_string(),
            });
        }
        if row == 0 || row > Self::MAX_ROW {
            return Err(XlexError::InvalidReference {
                reference: s.to_string(),
            });
        }

        Ok(Self { col, row })
    }

    /// Converts a column letter sequence to a number (A=1, B=2, ..., Z=26, AA=27).
    fn col_from_letters(s: &str) -> Option<u32> {
        let mut result: u32 = 0;
        for c in s.chars() {
            if !c.is_ascii_alphabetic() {
                return None;
            }
            let digit = (c.to_ascii_uppercase() as u32) - ('A' as u32) + 1;
            result = result.checked_mul(26)?.checked_add(digit)?;
        }
        if result == 0 || result > Self::MAX_COL {
            return None;
        }
        Some(result)
    }

    /// Public version of col_from_letters for CLI use.
    pub fn col_from_letters_pub(s: &str) -> Option<u32> {
        Self::col_from_letters(s)
    }

    /// Converts a column number to letters (1=A, 2=B, ..., 27=AA).
    pub fn col_to_letters(col: u32) -> String {
        let mut result = String::new();
        let mut n = col;
        while n > 0 {
            n -= 1;
            result.insert(0, (b'A' + (n % 26) as u8) as char);
            n /= 26;
        }
        result
    }

    /// Returns the A1-style string representation.
    pub fn to_a1(&self) -> String {
        format!("{}{}", Self::col_to_letters(self.col), self.row)
    }
}

impl fmt::Display for CellRef {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        write!(f, "{}", self.to_a1())
    }
}

impl FromStr for CellRef {
    type Err = XlexError;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        Self::parse(s)
    }
}

/// The value of a cell.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize, Default)]
#[serde(tag = "type", content = "value")]
pub enum CellValue {
    /// Empty cell
    #[default]
    Empty,
    /// String value
    String(String),
    /// Numeric value (includes integers and floats)
    Number(f64),
    /// Boolean value
    Boolean(bool),
    /// Formula (stores the formula string, not the calculated result)
    Formula {
        formula: String,
        /// Cached result, if available
        cached_result: Option<Box<CellValue>>,
    },
    /// Error value (e.g., #VALUE!, #REF!, #DIV/0!)
    Error(CellError),
    /// Date/time value (stored as Excel serial number)
    DateTime(f64),
}

impl CellValue {
    /// Creates a new string value.
    pub fn string(s: impl Into<String>) -> Self {
        Self::String(s.into())
    }

    /// Creates a new number value.
    pub fn number(n: f64) -> Self {
        Self::Number(n)
    }

    /// Creates a new boolean value.
    pub fn boolean(b: bool) -> Self {
        Self::Boolean(b)
    }

    /// Creates a new formula value.
    pub fn formula(f: impl Into<String>) -> Self {
        Self::Formula {
            formula: f.into(),
            cached_result: None,
        }
    }

    /// Returns true if this is an empty cell.
    pub fn is_empty(&self) -> bool {
        matches!(self, Self::Empty)
    }

    /// Returns the type name as a string.
    pub fn type_name(&self) -> &'static str {
        match self {
            Self::Empty => "empty",
            Self::String(_) => "string",
            Self::Number(_) => "number",
            Self::Boolean(_) => "boolean",
            Self::Formula { .. } => "formula",
            Self::Error(_) => "error",
            Self::DateTime(_) => "datetime",
        }
    }

    /// Tries to convert to a string representation.
    pub fn to_display_string(&self) -> String {
        match self {
            Self::Empty => String::new(),
            Self::String(s) => s.clone(),
            Self::Number(n) => {
                // Format nicely, avoiding unnecessary decimals
                if n.fract() == 0.0 && n.abs() < 1e15 {
                    format!("{:.0}", n)
                } else {
                    n.to_string()
                }
            }
            Self::Boolean(b) => if *b { "TRUE" } else { "FALSE" }.to_string(),
            Self::Formula { formula, .. } => format!("={}", formula),
            Self::Error(e) => e.to_string(),
            Self::DateTime(serial) => format!("{}", serial), // TODO: Format as date
        }
    }
}

impl fmt::Display for CellValue {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        write!(f, "{}", self.to_display_string())
    }
}

/// Excel cell error types.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum CellError {
    /// #NULL! - Intersection of two ranges that don't intersect
    Null,
    /// #DIV/0! - Division by zero
    DivZero,
    /// #VALUE! - Wrong type of argument or operand
    Value,
    /// #REF! - Invalid cell reference
    Ref,
    /// #NAME? - Unrecognized formula name
    Name,
    /// #NUM! - Invalid numeric value
    Num,
    /// #N/A - Value not available
    Na,
    /// #GETTING_DATA - Data retrieval in progress
    GettingData,
}

impl CellError {
    /// Parses an error string (e.g., "#VALUE!").
    pub fn parse(s: &str) -> Option<Self> {
        match s.to_uppercase().as_str() {
            "#NULL!" => Some(Self::Null),
            "#DIV/0!" => Some(Self::DivZero),
            "#VALUE!" => Some(Self::Value),
            "#REF!" => Some(Self::Ref),
            "#NAME?" => Some(Self::Name),
            "#NUM!" => Some(Self::Num),
            "#N/A" => Some(Self::Na),
            "#GETTING_DATA" => Some(Self::GettingData),
            _ => None,
        }
    }
}

impl fmt::Display for CellError {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        let s = match self {
            Self::Null => "#NULL!",
            Self::DivZero => "#DIV/0!",
            Self::Value => "#VALUE!",
            Self::Ref => "#REF!",
            Self::Name => "#NAME?",
            Self::Num => "#NUM!",
            Self::Na => "#N/A",
            Self::GettingData => "#GETTING_DATA",
        };
        write!(f, "{}", s)
    }
}

/// A cell with its reference, value, and optional style.
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct Cell {
    /// Cell reference
    pub reference: CellRef,
    /// Cell value
    pub value: CellValue,
    /// Style ID (index into the style registry)
    pub style_id: Option<u32>,
    /// Comment text, if any
    pub comment: Option<String>,
    /// Hyperlink URL, if any
    pub hyperlink: Option<String>,
}

impl Cell {
    /// Creates a new cell with the given reference and value.
    pub fn new(reference: CellRef, value: CellValue) -> Self {
        Self {
            reference,
            value,
            style_id: None,
            comment: None,
            hyperlink: None,
        }
    }

    /// Creates an empty cell at the given reference.
    pub fn empty(reference: CellRef) -> Self {
        Self::new(reference, CellValue::Empty)
    }

    /// Sets the style ID.
    pub fn with_style(mut self, style_id: u32) -> Self {
        self.style_id = Some(style_id);
        self
    }

    /// Sets a comment.
    pub fn with_comment(mut self, comment: impl Into<String>) -> Self {
        self.comment = Some(comment.into());
        self
    }

    /// Sets a hyperlink.
    pub fn with_hyperlink(mut self, url: impl Into<String>) -> Self {
        self.hyperlink = Some(url.into());
        self
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_cell_ref_parse() {
        assert_eq!(CellRef::parse("A1").unwrap(), CellRef::new(1, 1));
        assert_eq!(CellRef::parse("B2").unwrap(), CellRef::new(2, 2));
        assert_eq!(CellRef::parse("Z26").unwrap(), CellRef::new(26, 26));
        assert_eq!(CellRef::parse("AA1").unwrap(), CellRef::new(27, 1));
        assert_eq!(CellRef::parse("AA100").unwrap(), CellRef::new(27, 100));
        assert_eq!(
            CellRef::parse("XFD1048576").unwrap(),
            CellRef::new(16384, 1048576)
        );
    }

    #[test]
    fn test_cell_ref_parse_case_insensitive() {
        assert_eq!(CellRef::parse("a1").unwrap(), CellRef::new(1, 1));
        assert_eq!(CellRef::parse("Ab10").unwrap(), CellRef::new(28, 10));
    }

    #[test]
    fn test_cell_ref_parse_invalid() {
        assert!(CellRef::parse("").is_err());
        assert!(CellRef::parse("1A").is_err());
        assert!(CellRef::parse("A0").is_err());
        assert!(CellRef::parse("A").is_err());
        assert!(CellRef::parse("1").is_err());
        assert!(CellRef::parse("XFE1").is_err()); // Column too large
        assert!(CellRef::parse("A1048577").is_err()); // Row too large
    }

    #[test]
    fn test_cell_ref_to_a1() {
        assert_eq!(CellRef::new(1, 1).to_a1(), "A1");
        assert_eq!(CellRef::new(26, 26).to_a1(), "Z26");
        assert_eq!(CellRef::new(27, 1).to_a1(), "AA1");
        assert_eq!(CellRef::new(16384, 1048576).to_a1(), "XFD1048576");
    }

    #[test]
    fn test_col_to_letters() {
        assert_eq!(CellRef::col_to_letters(1), "A");
        assert_eq!(CellRef::col_to_letters(26), "Z");
        assert_eq!(CellRef::col_to_letters(27), "AA");
        assert_eq!(CellRef::col_to_letters(28), "AB");
        assert_eq!(CellRef::col_to_letters(702), "ZZ");
        assert_eq!(CellRef::col_to_letters(703), "AAA");
        assert_eq!(CellRef::col_to_letters(16384), "XFD");
    }

    #[test]
    fn test_cell_value_type_name() {
        assert_eq!(CellValue::Empty.type_name(), "empty");
        assert_eq!(CellValue::string("test").type_name(), "string");
        assert_eq!(CellValue::number(42.0).type_name(), "number");
        assert_eq!(CellValue::boolean(true).type_name(), "boolean");
        assert_eq!(CellValue::formula("A1+B1").type_name(), "formula");
    }

    #[test]
    fn test_cell_value_display() {
        assert_eq!(CellValue::Empty.to_display_string(), "");
        assert_eq!(CellValue::string("hello").to_display_string(), "hello");
        assert_eq!(CellValue::number(42.0).to_display_string(), "42");
        assert_eq!(CellValue::number(3.14).to_display_string(), "3.14");
        assert_eq!(CellValue::boolean(true).to_display_string(), "TRUE");
        assert_eq!(CellValue::boolean(false).to_display_string(), "FALSE");
        assert_eq!(CellValue::formula("A1+B1").to_display_string(), "=A1+B1");
    }

    #[test]
    fn test_cell_error_parse() {
        assert_eq!(CellError::parse("#VALUE!"), Some(CellError::Value));
        assert_eq!(CellError::parse("#REF!"), Some(CellError::Ref));
        assert_eq!(CellError::parse("#DIV/0!"), Some(CellError::DivZero));
        assert_eq!(CellError::parse("#N/A"), Some(CellError::Na));
        assert_eq!(CellError::parse("invalid"), None);
    }
}
