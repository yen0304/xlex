//! Sheet types and operations.

use std::fmt;

use serde::{Deserialize, Serialize};

use crate::cell::{Cell, CellRef, CellValue};

/// Sheet visibility state.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default, Serialize, Deserialize)]
pub enum SheetVisibility {
    /// Sheet is visible (default)
    #[default]
    Visible,
    /// Sheet is hidden (can be unhidden via UI)
    Hidden,
    /// Sheet is very hidden (cannot be unhidden via UI)
    VeryHidden,
}

impl fmt::Display for SheetVisibility {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::Visible => write!(f, "visible"),
            Self::Hidden => write!(f, "hidden"),
            Self::VeryHidden => write!(f, "veryHidden"),
        }
    }
}

impl SheetVisibility {
    /// Returns true if the sheet is visible.
    pub fn is_visible(&self) -> bool {
        matches!(self, Self::Visible)
    }

    /// Returns true if the sheet is hidden (either hidden or very hidden).
    pub fn is_hidden(&self) -> bool {
        !self.is_visible()
    }
}

/// Information about a sheet.
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SheetInfo {
    /// Sheet name
    pub name: String,
    /// Sheet ID (internal Excel ID)
    pub sheet_id: u32,
    /// Relationship ID (r:id)
    pub rel_id: String,
    /// Sheet visibility
    pub visibility: SheetVisibility,
    /// Index in the workbook (0-based)
    pub index: usize,
}

impl SheetInfo {
    /// Creates new sheet info.
    pub fn new(
        name: impl Into<String>,
        sheet_id: u32,
        rel_id: impl Into<String>,
        index: usize,
    ) -> Self {
        Self {
            name: name.into(),
            sheet_id,
            rel_id: rel_id.into(),
            visibility: SheetVisibility::Visible,
            index,
        }
    }
}

/// A worksheet containing cells.
#[derive(Debug, Clone)]
pub struct Sheet {
    /// Sheet information
    pub info: SheetInfo,
    /// Cells in the sheet (sparse storage)
    cells: std::collections::HashMap<(u32, u32), Cell>,
    /// Row heights (row -> height in points)
    row_heights: std::collections::HashMap<u32, f64>,
    /// Column widths (col -> width in characters)
    column_widths: std::collections::HashMap<u32, f64>,
    /// Hidden rows
    hidden_rows: std::collections::HashSet<u32>,
    /// Hidden columns
    hidden_columns: std::collections::HashSet<u32>,
    /// Merged cell ranges
    merged_ranges: Vec<crate::range::Range>,
    /// Used range (cached, may be None if not computed)
    used_range: Option<crate::range::Range>,
}

impl Sheet {
    /// Creates a new empty sheet.
    pub fn new(info: SheetInfo) -> Self {
        Self {
            info,
            cells: std::collections::HashMap::new(),
            row_heights: std::collections::HashMap::new(),
            column_widths: std::collections::HashMap::new(),
            hidden_rows: std::collections::HashSet::new(),
            hidden_columns: std::collections::HashSet::new(),
            merged_ranges: Vec::new(),
            used_range: None,
        }
    }

    /// Returns the sheet name.
    pub fn name(&self) -> &str {
        &self.info.name
    }

    /// Sets the sheet name.
    pub fn set_name(&mut self, name: impl Into<String>) {
        self.info.name = name.into();
    }

    /// Gets a cell by reference.
    pub fn get_cell(&self, cell_ref: &CellRef) -> Option<&Cell> {
        self.cells.get(&(cell_ref.col, cell_ref.row))
    }

    /// Gets a mutable cell by reference.
    pub fn get_cell_mut(&mut self, cell_ref: &CellRef) -> Option<&mut Cell> {
        self.cells.get_mut(&(cell_ref.col, cell_ref.row))
    }

    /// Gets the value of a cell.
    pub fn get_value(&self, cell_ref: &CellRef) -> CellValue {
        self.get_cell(cell_ref)
            .map(|c| c.value.clone())
            .unwrap_or(CellValue::Empty)
    }

    /// Sets a cell value.
    pub fn set_cell(&mut self, cell_ref: CellRef, value: CellValue) {
        let cell = Cell::new(cell_ref.clone(), value);
        self.cells.insert((cell_ref.col, cell_ref.row), cell);
        self.used_range = None; // Invalidate cache
    }

    /// Inserts a complete cell object, preserving all its properties (style_id, comment, hyperlink).
    pub fn insert_cell(&mut self, cell: Cell) {
        self.cells
            .insert((cell.reference.col, cell.reference.row), cell);
        self.used_range = None; // Invalidate cache
    }

    /// Sets a cell's style ID.
    pub fn set_cell_style(&mut self, cell_ref: &CellRef, style_id: Option<u32>) {
        if let Some(cell) = self.cells.get_mut(&(cell_ref.col, cell_ref.row)) {
            cell.style_id = style_id;
        } else if let Some(id) = style_id {
            // Create a new empty cell with the style
            let mut cell = Cell::empty(cell_ref.clone());
            cell.style_id = Some(id);
            self.cells.insert((cell_ref.col, cell_ref.row), cell);
        }
    }

    /// Sets a cell's comment.
    pub fn set_cell_comment(&mut self, cell_ref: &CellRef, comment: Option<String>) {
        if let Some(cell) = self.cells.get_mut(&(cell_ref.col, cell_ref.row)) {
            cell.comment = comment;
        } else if let Some(text) = comment {
            let mut cell = Cell::empty(cell_ref.clone());
            cell.comment = Some(text);
            self.cells.insert((cell_ref.col, cell_ref.row), cell);
        }
    }

    /// Sets a cell's hyperlink.
    pub fn set_cell_hyperlink(&mut self, cell_ref: &CellRef, hyperlink: Option<String>) {
        if let Some(cell) = self.cells.get_mut(&(cell_ref.col, cell_ref.row)) {
            cell.hyperlink = hyperlink;
        } else if let Some(url) = hyperlink {
            let mut cell = Cell::empty(cell_ref.clone());
            cell.hyperlink = Some(url);
            self.cells.insert((cell_ref.col, cell_ref.row), cell);
        }
    }

    /// Clears a cell.
    pub fn clear_cell(&mut self, cell_ref: &CellRef) {
        self.cells.remove(&(cell_ref.col, cell_ref.row));
        self.used_range = None;
    }

    /// Returns an iterator over all cells.
    pub fn cells(&self) -> impl Iterator<Item = &Cell> {
        self.cells.values()
    }

    /// Returns the number of non-empty cells.
    pub fn cell_count(&self) -> usize {
        self.cells.len()
    }

    /// Checks if the sheet is empty.
    pub fn is_empty(&self) -> bool {
        self.cells.is_empty()
    }

    /// Returns the dimensions of the sheet as (max_col, max_row).
    /// Returns (0, 0) if the sheet is empty.
    pub fn dimensions(&self) -> (u32, u32) {
        if self.cells.is_empty() {
            return (0, 0);
        }

        let mut max_col = 0;
        let mut max_row = 0;

        for (col, row) in self.cells.keys() {
            max_col = max_col.max(*col);
            max_row = max_row.max(*row);
        }

        (max_col, max_row)
    }

    /// Gets the row height.
    pub fn get_row_height(&self, row: u32) -> Option<f64> {
        self.row_heights.get(&row).copied()
    }

    /// Sets the row height.
    pub fn set_row_height(&mut self, row: u32, height: f64) {
        self.row_heights.insert(row, height);
    }

    /// Gets the column width.
    pub fn get_column_width(&self, col: u32) -> Option<f64> {
        self.column_widths.get(&col).copied()
    }

    /// Sets the column width.
    pub fn set_column_width(&mut self, col: u32, width: f64) {
        self.column_widths.insert(col, width);
    }

    /// Checks if a row is hidden.
    pub fn is_row_hidden(&self, row: u32) -> bool {
        self.hidden_rows.contains(&row)
    }

    /// Sets row visibility.
    pub fn set_row_hidden(&mut self, row: u32, hidden: bool) {
        if hidden {
            self.hidden_rows.insert(row);
        } else {
            self.hidden_rows.remove(&row);
        }
    }

    /// Checks if a column is hidden.
    pub fn is_column_hidden(&self, col: u32) -> bool {
        self.hidden_columns.contains(&col)
    }

    /// Sets column visibility.
    pub fn set_column_hidden(&mut self, col: u32, hidden: bool) {
        if hidden {
            self.hidden_columns.insert(col);
        } else {
            self.hidden_columns.remove(&col);
        }
    }

    /// Gets merged ranges.
    pub fn merged_ranges(&self) -> &[crate::range::Range] {
        &self.merged_ranges
    }

    /// Adds a merged range.
    pub fn add_merged_range(&mut self, range: crate::range::Range) {
        self.merged_ranges.push(range);
    }

    /// Removes a merged range.
    pub fn remove_merged_range(&mut self, range: &crate::range::Range) {
        self.merged_ranges.retain(|r| r != range);
    }

    /// Calculates and returns the used range.
    pub fn calculate_used_range(&mut self) -> Option<crate::range::Range> {
        if self.cells.is_empty() {
            self.used_range = None;
            return None;
        }

        let mut min_col = u32::MAX;
        let mut max_col = 0;
        let mut min_row = u32::MAX;
        let mut max_row = 0;

        for (col, row) in self.cells.keys() {
            min_col = min_col.min(*col);
            max_col = max_col.max(*col);
            min_row = min_row.min(*row);
            max_row = max_row.max(*row);
        }

        let range = crate::range::Range::new(
            CellRef::new(min_col, min_row),
            CellRef::new(max_col, max_row),
        );
        self.used_range = Some(range.clone());
        Some(range)
    }

    /// Returns the cached used range, or calculates it if not cached.
    pub fn used_range(&mut self) -> Option<crate::range::Range> {
        if self.used_range.is_none() {
            self.calculate_used_range();
        }
        self.used_range.clone()
    }

    /// Inserts a row at the specified position, shifting all rows below down by one.
    ///
    /// # Arguments
    /// * `row` - The row number (1-based) where the new row will be inserted
    /// * `count` - Number of rows to insert (default 1)
    pub fn insert_rows(&mut self, row: u32, count: u32) {
        if count == 0 {
            return;
        }

        // Collect all cells that need to be shifted
        let cells_to_shift: Vec<_> = self
            .cells
            .iter()
            .filter(|((_col, r), _)| *r >= row)
            .map(|((col, r), cell)| ((*col, *r), cell.clone()))
            .collect();

        // Remove old positions
        for ((col, r), _) in &cells_to_shift {
            self.cells.remove(&(*col, *r));
        }

        // Insert at new positions (shifted down)
        for ((col, r), cell) in cells_to_shift {
            self.cells.insert((col, r + count), cell);
        }

        // Shift row heights
        let heights_to_shift: Vec<_> = self
            .row_heights
            .iter()
            .filter(|(r, _)| **r >= row)
            .map(|(r, h)| (*r, *h))
            .collect();

        for (r, _) in &heights_to_shift {
            self.row_heights.remove(r);
        }

        for (r, h) in heights_to_shift {
            self.row_heights.insert(r + count, h);
        }

        // Shift hidden rows
        let hidden_to_shift: Vec<_> = self
            .hidden_rows
            .iter()
            .filter(|r| **r >= row)
            .copied()
            .collect();

        for r in &hidden_to_shift {
            self.hidden_rows.remove(r);
        }

        for r in hidden_to_shift {
            self.hidden_rows.insert(r + count);
        }

        // Shift merged ranges
        for range in &mut self.merged_ranges {
            if range.start.row >= row {
                range.start.row += count;
            }
            if range.end.row >= row {
                range.end.row += count;
            }
        }

        // Invalidate used range cache
        self.used_range = None;
    }

    /// Deletes rows starting at the specified position, shifting all rows below up.
    ///
    /// # Arguments
    /// * `row` - The row number (1-based) where deletion starts
    /// * `count` - Number of rows to delete (default 1)
    pub fn delete_rows(&mut self, row: u32, count: u32) {
        if count == 0 {
            return;
        }

        let end_row = row + count - 1;

        // Remove cells in the deleted rows
        self.cells.retain(|(_col, r), _| *r < row || *r > end_row);

        // Collect cells that need to be shifted up
        let cells_to_shift: Vec<_> = self
            .cells
            .iter()
            .filter(|((_col, r), _)| *r > end_row)
            .map(|((col, r), cell)| ((*col, *r), cell.clone()))
            .collect();

        // Remove old positions
        for ((col, r), _) in &cells_to_shift {
            self.cells.remove(&(*col, *r));
        }

        // Insert at new positions (shifted up)
        for ((col, r), cell) in cells_to_shift {
            self.cells.insert((col, r - count), cell);
        }

        // Remove deleted row heights and shift remaining
        for r in row..=end_row {
            self.row_heights.remove(&r);
        }

        let heights_to_shift: Vec<_> = self
            .row_heights
            .iter()
            .filter(|(r, _)| **r > end_row)
            .map(|(r, h)| (*r, *h))
            .collect();

        for (r, _) in &heights_to_shift {
            self.row_heights.remove(r);
        }

        for (r, h) in heights_to_shift {
            self.row_heights.insert(r - count, h);
        }

        // Remove deleted hidden rows and shift remaining
        for r in row..=end_row {
            self.hidden_rows.remove(&r);
        }

        let hidden_to_shift: Vec<_> = self
            .hidden_rows
            .iter()
            .filter(|r| **r > end_row)
            .copied()
            .collect();

        for r in &hidden_to_shift {
            self.hidden_rows.remove(r);
        }

        for r in hidden_to_shift {
            self.hidden_rows.insert(r - count);
        }

        // Update merged ranges (remove if fully deleted, adjust otherwise)
        self.merged_ranges.retain_mut(|range| {
            // Remove if fully within deleted range
            if range.start.row >= row && range.end.row <= end_row {
                return false;
            }

            // Adjust ranges
            if range.start.row > end_row {
                range.start.row -= count;
            } else if range.start.row >= row {
                range.start.row = row;
            }

            if range.end.row > end_row {
                range.end.row -= count;
            } else if range.end.row >= row {
                range.end.row = row.saturating_sub(1).max(range.start.row);
            }

            true
        });

        // Invalidate used range cache
        self.used_range = None;
    }

    /// Inserts columns at the specified position, shifting all columns to the right.
    ///
    /// # Arguments
    /// * `col` - The column number (1-based) where the new columns will be inserted
    /// * `count` - Number of columns to insert (default 1)
    pub fn insert_columns(&mut self, col: u32, count: u32) {
        if count == 0 {
            return;
        }

        // Collect all cells that need to be shifted
        let cells_to_shift: Vec<_> = self
            .cells
            .iter()
            .filter(|((c, _row), _)| *c >= col)
            .map(|((c, r), cell)| ((*c, *r), cell.clone()))
            .collect();

        // Remove old positions
        for ((c, r), _) in &cells_to_shift {
            self.cells.remove(&(*c, *r));
        }

        // Insert at new positions (shifted right)
        for ((c, r), cell) in cells_to_shift {
            self.cells.insert((c + count, r), cell);
        }

        // Shift column widths
        let widths_to_shift: Vec<_> = self
            .column_widths
            .iter()
            .filter(|(c, _)| **c >= col)
            .map(|(c, w)| (*c, *w))
            .collect();

        for (c, _) in &widths_to_shift {
            self.column_widths.remove(c);
        }

        for (c, w) in widths_to_shift {
            self.column_widths.insert(c + count, w);
        }

        // Shift hidden columns
        let hidden_to_shift: Vec<_> = self
            .hidden_columns
            .iter()
            .filter(|c| **c >= col)
            .copied()
            .collect();

        for c in &hidden_to_shift {
            self.hidden_columns.remove(c);
        }

        for c in hidden_to_shift {
            self.hidden_columns.insert(c + count);
        }

        // Shift merged ranges
        for range in &mut self.merged_ranges {
            if range.start.col >= col {
                range.start.col += count;
            }
            if range.end.col >= col {
                range.end.col += count;
            }
        }

        // Invalidate used range cache
        self.used_range = None;
    }

    /// Deletes columns starting at the specified position, shifting all columns to the left.
    ///
    /// # Arguments
    /// * `col` - The column number (1-based) where deletion starts
    /// * `count` - Number of columns to delete (default 1)
    pub fn delete_columns(&mut self, col: u32, count: u32) {
        if count == 0 {
            return;
        }

        let end_col = col + count - 1;

        // Remove cells in the deleted columns
        self.cells.retain(|(c, _row), _| *c < col || *c > end_col);

        // Collect cells that need to be shifted left
        let cells_to_shift: Vec<_> = self
            .cells
            .iter()
            .filter(|((c, _row), _)| *c > end_col)
            .map(|((c, r), cell)| ((*c, *r), cell.clone()))
            .collect();

        // Remove old positions
        for ((c, r), _) in &cells_to_shift {
            self.cells.remove(&(*c, *r));
        }

        // Insert at new positions (shifted left)
        for ((c, r), cell) in cells_to_shift {
            self.cells.insert((c - count, r), cell);
        }

        // Remove deleted column widths and shift remaining
        for c in col..=end_col {
            self.column_widths.remove(&c);
        }

        let widths_to_shift: Vec<_> = self
            .column_widths
            .iter()
            .filter(|(c, _)| **c > end_col)
            .map(|(c, w)| (*c, *w))
            .collect();

        for (c, _) in &widths_to_shift {
            self.column_widths.remove(c);
        }

        for (c, w) in widths_to_shift {
            self.column_widths.insert(c - count, w);
        }

        // Remove deleted hidden columns and shift remaining
        for c in col..=end_col {
            self.hidden_columns.remove(&c);
        }

        let hidden_to_shift: Vec<_> = self
            .hidden_columns
            .iter()
            .filter(|c| **c > end_col)
            .copied()
            .collect();

        for c in &hidden_to_shift {
            self.hidden_columns.remove(c);
        }

        for c in hidden_to_shift {
            self.hidden_columns.insert(c - count);
        }

        // Update merged ranges (remove if fully deleted, adjust otherwise)
        self.merged_ranges.retain_mut(|range| {
            // Remove if fully within deleted range
            if range.start.col >= col && range.end.col <= end_col {
                return false;
            }

            // Adjust ranges
            if range.start.col > end_col {
                range.start.col -= count;
            } else if range.start.col >= col {
                range.start.col = col;
            }

            if range.end.col > end_col {
                range.end.col -= count;
            } else if range.end.col >= col {
                range.end.col = col.saturating_sub(1).max(range.start.col);
            }

            true
        });

        // Invalidate used range cache
        self.used_range = None;
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn make_sheet() -> Sheet {
        let info = SheetInfo::new("Test", 1, "rId1", 0);
        Sheet::new(info)
    }

    #[test]
    fn test_sheet_basic() {
        let sheet = make_sheet();
        assert_eq!(sheet.name(), "Test");
        assert!(sheet.is_empty());
        assert_eq!(sheet.cell_count(), 0);
    }

    #[test]
    fn test_sheet_set_name() {
        let mut sheet = make_sheet();
        sheet.set_name("NewName");
        assert_eq!(sheet.name(), "NewName");
    }

    #[test]
    fn test_sheet_info() {
        let info = SheetInfo::new("MySheet", 5, "rId5", 4);
        assert_eq!(info.name, "MySheet");
        assert_eq!(info.sheet_id, 5);
        assert_eq!(info.rel_id, "rId5");
        assert_eq!(info.index, 4);
        assert_eq!(info.visibility, SheetVisibility::Visible);
    }

    #[test]
    fn test_sheet_visibility_enum() {
        assert!(SheetVisibility::Visible.is_visible());
        assert!(!SheetVisibility::Visible.is_hidden());

        assert!(!SheetVisibility::Hidden.is_visible());
        assert!(SheetVisibility::Hidden.is_hidden());

        assert!(!SheetVisibility::VeryHidden.is_visible());
        assert!(SheetVisibility::VeryHidden.is_hidden());
    }

    #[test]
    fn test_sheet_visibility_display() {
        assert_eq!(format!("{}", SheetVisibility::Visible), "visible");
        assert_eq!(format!("{}", SheetVisibility::Hidden), "hidden");
        assert_eq!(format!("{}", SheetVisibility::VeryHidden), "veryHidden");
    }

    #[test]
    fn test_sheet_visibility_default() {
        let visibility = SheetVisibility::default();
        assert_eq!(visibility, SheetVisibility::Visible);
    }

    #[test]
    fn test_sheet_set_cell() {
        let mut sheet = make_sheet();
        let cell_ref = CellRef::new(1, 1);

        sheet.set_cell(cell_ref.clone(), CellValue::string("Hello"));
        assert!(!sheet.is_empty());
        assert_eq!(sheet.cell_count(), 1);

        let value = sheet.get_value(&cell_ref);
        assert_eq!(value, CellValue::String("Hello".to_string()));
    }

    #[test]
    fn test_sheet_get_cell() {
        let mut sheet = make_sheet();
        let cell_ref = CellRef::new(2, 3);

        // Cell doesn't exist yet
        assert!(sheet.get_cell(&cell_ref).is_none());

        sheet.set_cell(cell_ref.clone(), CellValue::number(42.0));

        let cell = sheet.get_cell(&cell_ref).unwrap();
        assert_eq!(cell.reference, cell_ref);
        assert_eq!(cell.value, CellValue::number(42.0));
    }

    #[test]
    fn test_sheet_get_cell_mut() {
        let mut sheet = make_sheet();
        let cell_ref = CellRef::new(1, 1);

        sheet.set_cell(cell_ref.clone(), CellValue::string("original"));

        if let Some(cell) = sheet.get_cell_mut(&cell_ref) {
            cell.value = CellValue::string("modified");
        }

        let value = sheet.get_value(&cell_ref);
        assert_eq!(value, CellValue::string("modified"));
    }

    #[test]
    fn test_sheet_clear_cell() {
        let mut sheet = make_sheet();
        let cell_ref = CellRef::new(1, 1);

        sheet.set_cell(cell_ref.clone(), CellValue::string("Hello"));
        assert_eq!(sheet.cell_count(), 1);

        sheet.clear_cell(&cell_ref);
        assert!(sheet.is_empty());
    }

    #[test]
    fn test_sheet_row_column_dimensions() {
        let mut sheet = make_sheet();

        sheet.set_row_height(1, 25.0);
        assert_eq!(sheet.get_row_height(1), Some(25.0));
        assert_eq!(sheet.get_row_height(2), None);

        sheet.set_column_width(1, 15.0);
        assert_eq!(sheet.get_column_width(1), Some(15.0));
    }

    #[test]
    fn test_sheet_visibility() {
        let mut sheet = make_sheet();

        assert!(!sheet.is_row_hidden(1));
        sheet.set_row_hidden(1, true);
        assert!(sheet.is_row_hidden(1));
        sheet.set_row_hidden(1, false);
        assert!(!sheet.is_row_hidden(1));

        assert!(!sheet.is_column_hidden(1));
        sheet.set_column_hidden(1, true);
        assert!(sheet.is_column_hidden(1));
    }

    #[test]
    fn test_sheet_used_range() {
        let mut sheet = make_sheet();

        assert!(sheet.used_range().is_none());

        sheet.set_cell(CellRef::new(2, 3), CellValue::string("A"));
        sheet.set_cell(CellRef::new(5, 10), CellValue::string("B"));

        let range = sheet.used_range().unwrap();
        assert_eq!(range.start, CellRef::new(2, 3));
        assert_eq!(range.end, CellRef::new(5, 10));
    }

    #[test]
    fn test_sheet_insert_rows() {
        let mut sheet = make_sheet();

        // Set up initial data
        sheet.set_cell(CellRef::new(1, 1), CellValue::string("A1"));
        sheet.set_cell(CellRef::new(1, 2), CellValue::string("A2"));
        sheet.set_cell(CellRef::new(1, 3), CellValue::string("A3"));
        sheet.set_row_height(2, 25.0);
        sheet.set_row_hidden(3, true);

        // Insert 1 row at row 2
        sheet.insert_rows(2, 1);

        // Verify cells shifted
        assert_eq!(
            sheet.get_value(&CellRef::new(1, 1)),
            CellValue::string("A1")
        );
        assert_eq!(sheet.get_value(&CellRef::new(1, 2)), CellValue::Empty); // New empty row
        assert_eq!(
            sheet.get_value(&CellRef::new(1, 3)),
            CellValue::string("A2")
        );
        assert_eq!(
            sheet.get_value(&CellRef::new(1, 4)),
            CellValue::string("A3")
        );

        // Verify row properties shifted
        assert_eq!(sheet.get_row_height(3), Some(25.0));
        assert!(sheet.is_row_hidden(4));
    }

    #[test]
    fn test_sheet_delete_rows() {
        let mut sheet = make_sheet();

        // Set up initial data
        sheet.set_cell(CellRef::new(1, 1), CellValue::string("A1"));
        sheet.set_cell(CellRef::new(1, 2), CellValue::string("A2"));
        sheet.set_cell(CellRef::new(1, 3), CellValue::string("A3"));
        sheet.set_cell(CellRef::new(1, 4), CellValue::string("A4"));
        sheet.set_row_height(4, 30.0);

        // Delete row 2
        sheet.delete_rows(2, 1);

        // Verify cells shifted
        assert_eq!(
            sheet.get_value(&CellRef::new(1, 1)),
            CellValue::string("A1")
        );
        assert_eq!(
            sheet.get_value(&CellRef::new(1, 2)),
            CellValue::string("A3")
        );
        assert_eq!(
            sheet.get_value(&CellRef::new(1, 3)),
            CellValue::string("A4")
        );
        assert_eq!(sheet.cell_count(), 3);

        // Verify row height shifted
        assert_eq!(sheet.get_row_height(3), Some(30.0));
    }

    #[test]
    fn test_sheet_insert_columns() {
        let mut sheet = make_sheet();

        // Set up initial data
        sheet.set_cell(CellRef::new(1, 1), CellValue::string("A1"));
        sheet.set_cell(CellRef::new(2, 1), CellValue::string("B1"));
        sheet.set_cell(CellRef::new(3, 1), CellValue::string("C1"));
        sheet.set_column_width(2, 20.0);
        sheet.set_column_hidden(3, true);

        // Insert 1 column at column 2
        sheet.insert_columns(2, 1);

        // Verify cells shifted
        assert_eq!(
            sheet.get_value(&CellRef::new(1, 1)),
            CellValue::string("A1")
        );
        assert_eq!(sheet.get_value(&CellRef::new(2, 1)), CellValue::Empty); // New empty column
        assert_eq!(
            sheet.get_value(&CellRef::new(3, 1)),
            CellValue::string("B1")
        );
        assert_eq!(
            sheet.get_value(&CellRef::new(4, 1)),
            CellValue::string("C1")
        );

        // Verify column properties shifted
        assert_eq!(sheet.get_column_width(3), Some(20.0));
        assert!(sheet.is_column_hidden(4));
    }

    #[test]
    fn test_sheet_delete_columns() {
        let mut sheet = make_sheet();

        // Set up initial data
        sheet.set_cell(CellRef::new(1, 1), CellValue::string("A1"));
        sheet.set_cell(CellRef::new(2, 1), CellValue::string("B1"));
        sheet.set_cell(CellRef::new(3, 1), CellValue::string("C1"));
        sheet.set_cell(CellRef::new(4, 1), CellValue::string("D1"));
        sheet.set_column_width(4, 25.0);

        // Delete column 2
        sheet.delete_columns(2, 1);

        // Verify cells shifted
        assert_eq!(
            sheet.get_value(&CellRef::new(1, 1)),
            CellValue::string("A1")
        );
        assert_eq!(
            sheet.get_value(&CellRef::new(2, 1)),
            CellValue::string("C1")
        );
        assert_eq!(
            sheet.get_value(&CellRef::new(3, 1)),
            CellValue::string("D1")
        );
        assert_eq!(sheet.cell_count(), 3);

        // Verify column width shifted
        assert_eq!(sheet.get_column_width(3), Some(25.0));
    }

    #[test]
    fn test_sheet_insert_rows_zero_count() {
        let mut sheet = make_sheet();
        sheet.set_cell(CellRef::new(1, 1), CellValue::string("A1"));

        // Insert 0 rows should do nothing
        sheet.insert_rows(1, 0);

        assert_eq!(
            sheet.get_value(&CellRef::new(1, 1)),
            CellValue::string("A1")
        );
    }

    #[test]
    fn test_sheet_delete_rows_zero_count() {
        let mut sheet = make_sheet();
        sheet.set_cell(CellRef::new(1, 1), CellValue::string("A1"));

        // Delete 0 rows should do nothing
        sheet.delete_rows(1, 0);

        assert_eq!(
            sheet.get_value(&CellRef::new(1, 1)),
            CellValue::string("A1")
        );
    }

    #[test]
    fn test_sheet_insert_columns_zero_count() {
        let mut sheet = make_sheet();
        sheet.set_cell(CellRef::new(1, 1), CellValue::string("A1"));

        // Insert 0 columns should do nothing
        sheet.insert_columns(1, 0);

        assert_eq!(
            sheet.get_value(&CellRef::new(1, 1)),
            CellValue::string("A1")
        );
    }

    #[test]
    fn test_sheet_delete_columns_zero_count() {
        let mut sheet = make_sheet();
        sheet.set_cell(CellRef::new(1, 1), CellValue::string("A1"));

        // Delete 0 columns should do nothing
        sheet.delete_columns(1, 0);

        assert_eq!(
            sheet.get_value(&CellRef::new(1, 1)),
            CellValue::string("A1")
        );
    }

    #[test]
    fn test_sheet_insert_multiple_rows() {
        let mut sheet = make_sheet();
        sheet.set_cell(CellRef::new(1, 1), CellValue::string("A1"));
        sheet.set_cell(CellRef::new(1, 2), CellValue::string("A2"));

        // Insert 3 rows at row 2
        sheet.insert_rows(2, 3);

        assert_eq!(
            sheet.get_value(&CellRef::new(1, 1)),
            CellValue::string("A1")
        );
        assert_eq!(sheet.get_value(&CellRef::new(1, 2)), CellValue::Empty);
        assert_eq!(sheet.get_value(&CellRef::new(1, 3)), CellValue::Empty);
        assert_eq!(sheet.get_value(&CellRef::new(1, 4)), CellValue::Empty);
        assert_eq!(
            sheet.get_value(&CellRef::new(1, 5)),
            CellValue::string("A2")
        );
    }

    #[test]
    fn test_sheet_delete_multiple_rows() {
        let mut sheet = make_sheet();
        for i in 1..=5 {
            sheet.set_cell(CellRef::new(1, i), CellValue::number(i as f64));
        }

        // Delete rows 2-4 (3 rows)
        sheet.delete_rows(2, 3);

        assert_eq!(sheet.get_value(&CellRef::new(1, 1)), CellValue::number(1.0));
        assert_eq!(sheet.get_value(&CellRef::new(1, 2)), CellValue::number(5.0));
        assert_eq!(sheet.cell_count(), 2);
    }

    #[test]
    fn test_sheet_cells_iterator() {
        let mut sheet = make_sheet();
        sheet.set_cell(CellRef::new(1, 1), CellValue::string("A"));
        sheet.set_cell(CellRef::new(2, 2), CellValue::string("B"));
        sheet.set_cell(CellRef::new(3, 3), CellValue::string("C"));

        let cells: Vec<_> = sheet.cells().collect();
        assert_eq!(cells.len(), 3);
    }

    #[test]
    fn test_sheet_dimensions() {
        let mut sheet = make_sheet();

        // Empty sheet
        assert_eq!(sheet.dimensions(), (0, 0));

        sheet.set_cell(CellRef::new(5, 10), CellValue::string("A"));
        assert_eq!(sheet.dimensions(), (5, 10));

        sheet.set_cell(CellRef::new(10, 5), CellValue::string("B"));
        assert_eq!(sheet.dimensions(), (10, 10));
    }

    #[test]
    fn test_sheet_set_cell_style() {
        let mut sheet = make_sheet();
        let cell_ref = CellRef::new(1, 1);

        // Set style on non-existent cell creates empty cell with style
        sheet.set_cell_style(&cell_ref, Some(5));
        let cell = sheet.get_cell(&cell_ref).unwrap();
        assert_eq!(cell.style_id, Some(5));
        assert!(cell.value.is_empty());

        // Set style on existing cell
        sheet.set_cell(CellRef::new(2, 2), CellValue::string("test"));
        sheet.set_cell_style(&CellRef::new(2, 2), Some(10));
        let cell = sheet.get_cell(&CellRef::new(2, 2)).unwrap();
        assert_eq!(cell.style_id, Some(10));
        assert_eq!(cell.value, CellValue::string("test"));

        // Remove style
        sheet.set_cell_style(&CellRef::new(2, 2), None);
        let cell = sheet.get_cell(&CellRef::new(2, 2)).unwrap();
        assert_eq!(cell.style_id, None);
    }

    #[test]
    fn test_sheet_set_cell_comment() {
        let mut sheet = make_sheet();
        let cell_ref = CellRef::new(1, 1);

        // Set comment on non-existent cell creates empty cell with comment
        sheet.set_cell_comment(&cell_ref, Some("A comment".to_string()));
        let cell = sheet.get_cell(&cell_ref).unwrap();
        assert_eq!(cell.comment, Some("A comment".to_string()));
        assert!(cell.value.is_empty());

        // Set comment on existing cell
        sheet.set_cell(CellRef::new(2, 2), CellValue::string("test"));
        sheet.set_cell_comment(&CellRef::new(2, 2), Some("Another comment".to_string()));
        let cell = sheet.get_cell(&CellRef::new(2, 2)).unwrap();
        assert_eq!(cell.comment, Some("Another comment".to_string()));

        // Remove comment
        sheet.set_cell_comment(&CellRef::new(2, 2), None);
        let cell = sheet.get_cell(&CellRef::new(2, 2)).unwrap();
        assert_eq!(cell.comment, None);
    }

    #[test]
    fn test_sheet_set_cell_hyperlink() {
        let mut sheet = make_sheet();
        let cell_ref = CellRef::new(1, 1);

        // Set hyperlink on non-existent cell creates empty cell with hyperlink
        sheet.set_cell_hyperlink(&cell_ref, Some("https://example.com".to_string()));
        let cell = sheet.get_cell(&cell_ref).unwrap();
        assert_eq!(cell.hyperlink, Some("https://example.com".to_string()));
        assert!(cell.value.is_empty());

        // Set hyperlink on existing cell
        sheet.set_cell(CellRef::new(2, 2), CellValue::string("Click"));
        sheet.set_cell_hyperlink(&CellRef::new(2, 2), Some("https://test.com".to_string()));
        let cell = sheet.get_cell(&CellRef::new(2, 2)).unwrap();
        assert_eq!(cell.hyperlink, Some("https://test.com".to_string()));

        // Remove hyperlink
        sheet.set_cell_hyperlink(&CellRef::new(2, 2), None);
        let cell = sheet.get_cell(&CellRef::new(2, 2)).unwrap();
        assert_eq!(cell.hyperlink, None);
    }

    #[test]
    fn test_sheet_merged_ranges() {
        let mut sheet = make_sheet();

        assert!(sheet.merged_ranges().is_empty());

        let range1 = crate::range::Range::parse("A1:B2").unwrap();
        let range2 = crate::range::Range::parse("C3:D4").unwrap();

        sheet.add_merged_range(range1.clone());
        sheet.add_merged_range(range2.clone());

        assert_eq!(sheet.merged_ranges().len(), 2);

        sheet.remove_merged_range(&range1);
        assert_eq!(sheet.merged_ranges().len(), 1);
        assert_eq!(sheet.merged_ranges()[0], range2);
    }

    #[test]
    fn test_sheet_calculate_used_range() {
        let mut sheet = make_sheet();

        // Empty sheet has no used range
        assert!(sheet.calculate_used_range().is_none());

        sheet.set_cell(CellRef::new(2, 3), CellValue::string("A"));
        sheet.set_cell(CellRef::new(5, 8), CellValue::string("B"));
        sheet.set_cell(CellRef::new(1, 10), CellValue::string("C"));

        let range = sheet.calculate_used_range().unwrap();
        assert_eq!(range.start, CellRef::new(1, 3));
        assert_eq!(range.end, CellRef::new(5, 10));
    }

    #[test]
    fn test_sheet_delete_rows_with_merged_ranges() {
        let mut sheet = make_sheet();

        // Add a merged range
        let range = crate::range::Range::parse("A2:B4").unwrap();
        sheet.add_merged_range(range);

        // Delete row 3 (inside the merged range)
        sheet.delete_rows(3, 1);

        // Merged range should be adjusted
        let ranges = sheet.merged_ranges();
        assert_eq!(ranges.len(), 1);
        // The range should be adjusted
    }

    #[test]
    fn test_sheet_insert_rows_with_merged_ranges() {
        let mut sheet = make_sheet();

        // Add a merged range starting at row 2
        let range = crate::range::Range::parse("A2:B4").unwrap();
        sheet.add_merged_range(range);

        // Insert 2 rows at row 2
        sheet.insert_rows(2, 2);

        // Merged range should be shifted
        let ranges = sheet.merged_ranges();
        assert_eq!(ranges.len(), 1);
        assert_eq!(ranges[0].start.row, 4);
        assert_eq!(ranges[0].end.row, 6);
    }

    #[test]
    fn test_sheet_delete_hidden_rows() {
        let mut sheet = make_sheet();

        sheet.set_row_hidden(2, true);
        sheet.set_row_hidden(3, true);
        sheet.set_row_hidden(5, true);

        // Delete row 3
        sheet.delete_rows(3, 1);

        // Row 2 should still be hidden
        assert!(sheet.is_row_hidden(2));
        // Row 3 (was 5) should be hidden now
        // Row 4 is now what was row 5
        assert!(sheet.is_row_hidden(4));
        // Old row 3 should be gone
        assert!(!sheet.is_row_hidden(3));
    }

    #[test]
    fn test_sheet_delete_columns_with_merged_ranges() {
        let mut sheet = make_sheet();

        // Add a merged range
        let range = crate::range::Range::parse("B1:D3").unwrap();
        sheet.add_merged_range(range);

        // Delete column C (column 3)
        sheet.delete_columns(3, 1);

        // Merged range should be adjusted
        let ranges = sheet.merged_ranges();
        assert_eq!(ranges.len(), 1);
    }
}
