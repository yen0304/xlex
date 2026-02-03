//! Range types and operations.

use serde::{Deserialize, Serialize};
use std::fmt;
use std::str::FromStr;

use crate::cell::CellRef;
use crate::error::{XlexError, XlexResult};

/// A range of cells (e.g., A1:B10).
#[derive(Debug, Clone, PartialEq, Eq, Hash, Serialize, Deserialize)]
pub struct Range {
    /// Start cell (top-left)
    pub start: CellRef,
    /// End cell (bottom-right)
    pub end: CellRef,
}

impl Range {
    /// Creates a new range from start to end.
    pub fn new(start: CellRef, end: CellRef) -> Self {
        Self { start, end }
    }

    /// Creates a range from a single cell.
    pub fn single(cell: CellRef) -> Self {
        Self {
            start: cell.clone(),
            end: cell,
        }
    }

    /// Parses a range string (e.g., "A1:B10", "A1", "A:A", "1:1").
    pub fn parse(s: &str) -> XlexResult<Self> {
        let s = s.trim();
        if s.is_empty() {
            return Err(XlexError::InvalidRange {
                range: s.to_string(),
            });
        }

        // Check for colon separator
        if let Some(colon_pos) = s.find(':') {
            let start_str = &s[..colon_pos];
            let end_str = &s[colon_pos + 1..];

            // Handle full column ranges (A:A, A:B)
            if start_str.chars().all(|c| c.is_ascii_alphabetic())
                && end_str.chars().all(|c| c.is_ascii_alphabetic())
            {
                let start_col = CellRef::col_from_letters_pub(start_str).ok_or_else(|| {
                    XlexError::InvalidRange {
                        range: s.to_string(),
                    }
                })?;
                let end_col = CellRef::col_from_letters_pub(end_str).ok_or_else(|| {
                    XlexError::InvalidRange {
                        range: s.to_string(),
                    }
                })?;

                if start_col > end_col {
                    return Err(XlexError::InvalidRange {
                        range: s.to_string(),
                    });
                }

                return Ok(Self {
                    start: CellRef::new(start_col, 1),
                    end: CellRef::new(end_col, CellRef::MAX_ROW),
                });
            }

            // Handle full row ranges (1:1, 1:10)
            if start_str.chars().all(|c| c.is_ascii_digit())
                && end_str.chars().all(|c| c.is_ascii_digit())
            {
                let start_row: u32 = start_str.parse().map_err(|_| XlexError::InvalidRange {
                    range: s.to_string(),
                })?;
                let end_row: u32 = end_str.parse().map_err(|_| XlexError::InvalidRange {
                    range: s.to_string(),
                })?;

                if start_row == 0 || end_row == 0 || start_row > end_row {
                    return Err(XlexError::InvalidRange {
                        range: s.to_string(),
                    });
                }

                return Ok(Self {
                    start: CellRef::new(1, start_row),
                    end: CellRef::new(CellRef::MAX_COL, end_row),
                });
            }

            // Normal cell range
            let start = CellRef::parse(start_str)?;
            let end = CellRef::parse(end_str)?;

            // Validate that end is not before start
            if end.col < start.col || end.row < start.row {
                return Err(XlexError::InvalidRange {
                    range: s.to_string(),
                });
            }

            Ok(Self { start, end })
        } else {
            // Single cell
            let cell = CellRef::parse(s)?;
            Ok(Self::single(cell))
        }
    }

    /// Returns the number of columns in this range.
    pub fn width(&self) -> u32 {
        self.end.col - self.start.col + 1
    }

    /// Returns the number of rows in this range.
    pub fn height(&self) -> u32 {
        self.end.row - self.start.row + 1
    }

    /// Returns the total number of cells in this range.
    pub fn cell_count(&self) -> u64 {
        self.width() as u64 * self.height() as u64
    }

    /// Returns true if this range contains a single cell.
    pub fn is_single(&self) -> bool {
        self.start == self.end
    }

    /// Returns true if this range contains the given cell reference.
    pub fn contains(&self, cell: &CellRef) -> bool {
        cell.col >= self.start.col
            && cell.col <= self.end.col
            && cell.row >= self.start.row
            && cell.row <= self.end.row
    }

    /// Returns an iterator over all cell references in this range.
    pub fn cells(&self) -> RangeCellIterator {
        RangeCellIterator {
            range: self.clone(),
            current_col: self.start.col,
            current_row: self.start.row,
        }
    }

    /// Returns the A1-style string representation.
    pub fn to_a1(&self) -> String {
        if self.is_single() {
            self.start.to_a1()
        } else {
            format!("{}:{}", self.start.to_a1(), self.end.to_a1())
        }
    }
}

impl fmt::Display for Range {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        write!(f, "{}", self.to_a1())
    }
}

impl FromStr for Range {
    type Err = XlexError;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        Self::parse(s)
    }
}

/// Iterator over cell references in a range.
pub struct RangeCellIterator {
    range: Range,
    current_col: u32,
    current_row: u32,
}

impl Iterator for RangeCellIterator {
    type Item = CellRef;

    fn next(&mut self) -> Option<Self::Item> {
        if self.current_row > self.range.end.row {
            return None;
        }

        let cell = CellRef::new(self.current_col, self.current_row);

        // Move to next cell
        self.current_col += 1;
        if self.current_col > self.range.end.col {
            self.current_col = self.range.start.col;
            self.current_row += 1;
        }

        Some(cell)
    }

    fn size_hint(&self) -> (usize, Option<usize>) {
        let remaining = self.range.cell_count() as usize;
        (remaining, Some(remaining))
    }
}

impl ExactSizeIterator for RangeCellIterator {}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_range_parse_normal() {
        let range = Range::parse("A1:B10").unwrap();
        assert_eq!(range.start, CellRef::new(1, 1));
        assert_eq!(range.end, CellRef::new(2, 10));
    }

    #[test]
    fn test_range_parse_single() {
        let range = Range::parse("A1").unwrap();
        assert!(range.is_single());
        assert_eq!(range.start, CellRef::new(1, 1));
        assert_eq!(range.end, CellRef::new(1, 1));
    }

    #[test]
    fn test_range_parse_full_column() {
        let range = Range::parse("A:A").unwrap();
        assert_eq!(range.start.col, 1);
        assert_eq!(range.end.col, 1);
        assert_eq!(range.start.row, 1);
        assert_eq!(range.end.row, CellRef::MAX_ROW);
    }

    #[test]
    fn test_range_parse_full_column_range() {
        let range = Range::parse("A:C").unwrap();
        assert_eq!(range.start.col, 1);
        assert_eq!(range.end.col, 3);
        assert_eq!(range.start.row, 1);
        assert_eq!(range.end.row, CellRef::MAX_ROW);
    }

    #[test]
    fn test_range_parse_full_row() {
        let range = Range::parse("1:1").unwrap();
        assert_eq!(range.start.col, 1);
        assert_eq!(range.end.col, CellRef::MAX_COL);
        assert_eq!(range.start.row, 1);
        assert_eq!(range.end.row, 1);
    }

    #[test]
    fn test_range_parse_full_row_range() {
        let range = Range::parse("1:10").unwrap();
        assert_eq!(range.start.col, 1);
        assert_eq!(range.end.col, CellRef::MAX_COL);
        assert_eq!(range.start.row, 1);
        assert_eq!(range.end.row, 10);
    }

    #[test]
    fn test_range_parse_invalid() {
        assert!(Range::parse("").is_err());
        assert!(Range::parse("B1:A1").is_err()); // End before start
        assert!(Range::parse("A10:A1").is_err()); // End row before start row
        assert!(Range::parse("C:A").is_err()); // End column before start column
        assert!(Range::parse("10:1").is_err()); // End row before start row (full row)
        assert!(Range::parse("0:1").is_err()); // Row 0 is invalid
    }

    #[test]
    fn test_range_dimensions() {
        let range = Range::parse("A1:C5").unwrap();
        assert_eq!(range.width(), 3);
        assert_eq!(range.height(), 5);
        assert_eq!(range.cell_count(), 15);
    }

    #[test]
    fn test_range_dimensions_single_cell() {
        let range = Range::single(CellRef::new(5, 5));
        assert_eq!(range.width(), 1);
        assert_eq!(range.height(), 1);
        assert_eq!(range.cell_count(), 1);
    }

    #[test]
    fn test_range_dimensions_large() {
        let range = Range::parse("A1:Z100").unwrap();
        assert_eq!(range.width(), 26);
        assert_eq!(range.height(), 100);
        assert_eq!(range.cell_count(), 2600);
    }

    #[test]
    fn test_range_contains() {
        let range = Range::parse("B2:D4").unwrap();
        assert!(range.contains(&CellRef::new(2, 2))); // Start
        assert!(range.contains(&CellRef::new(3, 3))); // Middle
        assert!(range.contains(&CellRef::new(4, 4))); // End
        assert!(!range.contains(&CellRef::new(1, 1))); // Before start
        assert!(!range.contains(&CellRef::new(5, 5))); // After end
        assert!(!range.contains(&CellRef::new(1, 3))); // Left of range
        assert!(!range.contains(&CellRef::new(5, 3))); // Right of range
        assert!(!range.contains(&CellRef::new(3, 1))); // Above range
        assert!(!range.contains(&CellRef::new(3, 5))); // Below range
    }

    #[test]
    fn test_range_contains_single_cell() {
        let range = Range::single(CellRef::new(3, 3));
        assert!(range.contains(&CellRef::new(3, 3)));
        assert!(!range.contains(&CellRef::new(2, 3)));
        assert!(!range.contains(&CellRef::new(4, 3)));
        assert!(!range.contains(&CellRef::new(3, 2)));
        assert!(!range.contains(&CellRef::new(3, 4)));
    }

    #[test]
    fn test_range_iterator() {
        let range = Range::parse("A1:B2").unwrap();
        let cells: Vec<CellRef> = range.cells().collect();
        assert_eq!(
            cells,
            vec![
                CellRef::new(1, 1),
                CellRef::new(2, 1),
                CellRef::new(1, 2),
                CellRef::new(2, 2),
            ]
        );
    }

    #[test]
    fn test_range_iterator_single() {
        let range = Range::single(CellRef::new(5, 5));
        let cells: Vec<CellRef> = range.cells().collect();
        assert_eq!(cells, vec![CellRef::new(5, 5)]);
    }

    #[test]
    fn test_range_iterator_row() {
        let range = Range::parse("A1:C1").unwrap();
        let cells: Vec<CellRef> = range.cells().collect();
        assert_eq!(
            cells,
            vec![CellRef::new(1, 1), CellRef::new(2, 1), CellRef::new(3, 1)]
        );
    }

    #[test]
    fn test_range_iterator_column() {
        let range = Range::parse("A1:A3").unwrap();
        let cells: Vec<CellRef> = range.cells().collect();
        assert_eq!(
            cells,
            vec![CellRef::new(1, 1), CellRef::new(1, 2), CellRef::new(1, 3)]
        );
    }

    #[test]
    fn test_range_iterator_size_hint() {
        let range = Range::parse("A1:C3").unwrap();
        let iter = range.cells();
        assert_eq!(iter.size_hint(), (9, Some(9)));
        assert_eq!(iter.len(), 9);
    }

    #[test]
    fn test_range_to_a1() {
        assert_eq!(Range::parse("A1:B10").unwrap().to_a1(), "A1:B10");
        assert_eq!(Range::parse("A1").unwrap().to_a1(), "A1");
        assert_eq!(Range::parse("AA1:ZZ100").unwrap().to_a1(), "AA1:ZZ100");
    }

    #[test]
    fn test_range_display() {
        let range = Range::parse("A1:B10").unwrap();
        assert_eq!(format!("{}", range), "A1:B10");

        let single = Range::single(CellRef::new(3, 5));
        assert_eq!(format!("{}", single), "C5");
    }

    #[test]
    fn test_range_from_str() {
        let range: Range = "A1:B10".parse().unwrap();
        assert_eq!(range.start, CellRef::new(1, 1));
        assert_eq!(range.end, CellRef::new(2, 10));

        // Invalid should fail
        assert!("invalid".parse::<Range>().is_err());
    }

    #[test]
    fn test_range_new() {
        let range = Range::new(CellRef::new(1, 1), CellRef::new(3, 5));
        assert_eq!(range.start, CellRef::new(1, 1));
        assert_eq!(range.end, CellRef::new(3, 5));
    }

    #[test]
    fn test_range_single() {
        let cell = CellRef::new(5, 10);
        let range = Range::single(cell.clone());
        assert_eq!(range.start, cell);
        assert_eq!(range.end, cell);
        assert!(range.is_single());
    }

    #[test]
    fn test_range_equality_and_hash() {
        use std::collections::HashSet;

        let range1 = Range::parse("A1:B10").unwrap();
        let range2 = Range::parse("A1:B10").unwrap();
        let range3 = Range::parse("A1:C10").unwrap();

        assert_eq!(range1, range2);
        assert_ne!(range1, range3);

        let mut set = HashSet::new();
        set.insert(range1.clone());
        assert!(set.contains(&range2));
        assert!(!set.contains(&range3));
    }

    #[test]
    fn test_range_parse_whitespace() {
        let range = Range::parse("  A1:B10  ").unwrap();
        assert_eq!(range.start, CellRef::new(1, 1));
        assert_eq!(range.end, CellRef::new(2, 10));
    }

    #[test]
    fn test_range_full_column_invalid() {
        // Column range with end before start
        assert!(Range::parse("Z:A").is_err());
    }

    #[test]
    fn test_range_full_row_zero() {
        // Row 0 is invalid
        assert!(Range::parse("0:10").is_err());
    }
}
