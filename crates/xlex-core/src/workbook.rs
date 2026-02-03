//! Workbook type and operations.

use std::collections::HashMap;
use std::io::{Read, Seek};
use std::path::{Path, PathBuf};

use chrono::{DateTime, Utc};
use serde::{Deserialize, Serialize};
use zip::ZipArchive;

use crate::cell::{CellRef, CellValue};
use crate::error::{XlexError, XlexResult};
use crate::parser::WorkbookParser;
use crate::sheet::{Sheet, SheetInfo, SheetVisibility};
use crate::style::StyleRegistry;

/// Document properties.
#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct DocumentProperties {
    pub title: Option<String>,
    pub subject: Option<String>,
    pub creator: Option<String>,
    pub keywords: Option<String>,
    pub description: Option<String>,
    pub last_modified_by: Option<String>,
    pub created: Option<DateTime<Utc>>,
    pub modified: Option<DateTime<Utc>>,
    pub category: Option<String>,
    pub content_status: Option<String>,
}

/// Workbook statistics.
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct WorkbookStats {
    /// Total number of sheets
    pub sheet_count: usize,
    /// Total number of cells across all sheets
    pub total_cells: usize,
    /// Total number of formulas
    pub formula_count: usize,
    /// Total number of styles
    pub style_count: usize,
    /// Total number of unique strings
    pub string_count: usize,
    /// File size in bytes
    pub file_size: u64,
}

/// A named range definition.
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct DefinedName {
    /// The name of the defined name
    pub name: String,
    /// The reference formula (e.g., "Sheet1!$A$1:$B$10")
    pub reference: String,
    /// Sheet scope (None for global scope, sheet index for local scope)
    pub local_sheet_id: Option<usize>,
    /// Comment/description
    pub comment: Option<String>,
    /// Whether this is hidden
    pub hidden: bool,
}

impl DefinedName {
    /// Creates a new defined name with global scope.
    pub fn new(name: impl Into<String>, reference: impl Into<String>) -> Self {
        Self {
            name: name.into(),
            reference: reference.into(),
            local_sheet_id: None,
            comment: None,
            hidden: false,
        }
    }

    /// Creates a new defined name with sheet scope.
    pub fn with_sheet_scope(
        name: impl Into<String>,
        reference: impl Into<String>,
        sheet_id: usize,
    ) -> Self {
        Self {
            name: name.into(),
            reference: reference.into(),
            local_sheet_id: Some(sheet_id),
            comment: None,
            hidden: false,
        }
    }
}

/// An Excel workbook.
#[derive(Debug)]
pub struct Workbook {
    /// File path (None for new workbooks)
    path: Option<PathBuf>,
    /// Document properties
    properties: DocumentProperties,
    /// Sheets in order
    sheets: Vec<Sheet>,
    /// Sheet name to index mapping
    sheet_map: HashMap<String, usize>,
    /// Style registry
    style_registry: StyleRegistry,
    /// Shared strings
    shared_strings: Vec<String>,
    /// Defined names (named ranges)
    defined_names: Vec<DefinedName>,
    /// Active sheet index
    active_sheet: usize,
    /// Modified flag
    modified: bool,
}

impl Workbook {
    /// Opens an existing workbook from a file.
    ///
    /// Automatically uses memory mapping for large files (>10MB) for better performance.
    pub fn open(path: impl AsRef<Path>) -> XlexResult<Self> {
        let path = path.as_ref();

        // Check extension
        if path.extension().and_then(|e| e.to_str()) != Some("xlsx") {
            return Err(XlexError::InvalidExtension {
                path: path.to_path_buf(),
            });
        }

        // Use WorkbookReader for automatic mmap handling
        let wb_reader = crate::reader::WorkbookReader::open(path)?;
        let cursor = std::io::Cursor::new(wb_reader.as_bytes());
        Self::from_reader(cursor, Some(path.to_path_buf()))
    }

    /// Creates a workbook from a reader.
    pub fn from_reader<R: Read + Seek>(reader: R, path: Option<PathBuf>) -> XlexResult<Self> {
        let mut archive = ZipArchive::new(reader)?;
        let parser = WorkbookParser::new();
        parser.parse(&mut archive, path)
    }

    /// Creates a new empty workbook.
    pub fn new() -> Self {
        let mut workbook = Self {
            path: None,
            properties: DocumentProperties::default(),
            sheets: Vec::new(),
            sheet_map: HashMap::new(),
            style_registry: StyleRegistry::new(),
            shared_strings: Vec::new(),
            defined_names: Vec::new(),
            active_sheet: 0,
            modified: true,
        };

        // Add default sheet
        workbook.add_sheet_internal("Sheet1");
        workbook
    }

    /// Creates a new workbook with specified sheet names.
    pub fn with_sheets(sheet_names: &[&str]) -> Self {
        let mut workbook = Self {
            path: None,
            properties: DocumentProperties::default(),
            sheets: Vec::new(),
            sheet_map: HashMap::new(),
            style_registry: StyleRegistry::new(),
            shared_strings: Vec::new(),
            defined_names: Vec::new(),
            active_sheet: 0,
            modified: true,
        };

        for name in sheet_names {
            workbook.add_sheet_internal(name);
        }

        // Ensure at least one sheet
        if workbook.sheets.is_empty() {
            workbook.add_sheet_internal("Sheet1");
        }

        workbook
    }

    fn add_sheet_internal(&mut self, name: &str) -> usize {
        let index = self.sheets.len();
        let sheet_id = (index + 1) as u32;
        let rel_id = format!("rId{}", sheet_id);

        let info = SheetInfo::new(name, sheet_id, rel_id, index);
        let sheet = Sheet::new(info);

        self.sheet_map.insert(name.to_string(), index);
        self.sheets.push(sheet);
        self.modified = true;

        index
    }

    /// Returns the file path if the workbook was opened from a file.
    pub fn path(&self) -> Option<&Path> {
        self.path.as_deref()
    }

    /// Returns the document properties.
    pub fn properties(&self) -> &DocumentProperties {
        &self.properties
    }

    /// Returns a mutable reference to document properties.
    pub fn properties_mut(&mut self) -> &mut DocumentProperties {
        self.modified = true;
        &mut self.properties
    }

    /// Returns the number of sheets.
    pub fn sheet_count(&self) -> usize {
        self.sheets.len()
    }

    /// Returns the names of all sheets.
    pub fn sheet_names(&self) -> Vec<&str> {
        self.sheets.iter().map(|s| s.name()).collect()
    }

    /// Gets a sheet by name.
    pub fn get_sheet(&self, name: &str) -> Option<&Sheet> {
        self.sheet_map.get(name).map(|&idx| &self.sheets[idx])
    }

    /// Gets a mutable sheet by name.
    pub fn get_sheet_mut(&mut self, name: &str) -> Option<&mut Sheet> {
        if let Some(&idx) = self.sheet_map.get(name) {
            self.modified = true;
            Some(&mut self.sheets[idx])
        } else {
            None
        }
    }

    /// Gets a sheet by index.
    pub fn get_sheet_by_index(&self, index: usize) -> Option<&Sheet> {
        self.sheets.get(index)
    }

    /// Gets a mutable sheet by index.
    pub fn get_sheet_by_index_mut(&mut self, index: usize) -> Option<&mut Sheet> {
        if index < self.sheets.len() {
            self.modified = true;
            Some(&mut self.sheets[index])
        } else {
            None
        }
    }

    /// Adds a new sheet with the given name.
    pub fn add_sheet(&mut self, name: &str) -> XlexResult<usize> {
        // Validate name
        Self::validate_sheet_name(name)?;

        // Check for duplicate
        if self.sheet_map.contains_key(name) {
            return Err(XlexError::SheetAlreadyExists {
                name: name.to_string(),
            });
        }

        Ok(self.add_sheet_internal(name))
    }

    /// Removes a sheet by name.
    pub fn remove_sheet(&mut self, name: &str) -> XlexResult<()> {
        // Can't delete the last sheet
        if self.sheets.len() <= 1 {
            return Err(XlexError::CannotDeleteLastSheet);
        }

        let index = *self
            .sheet_map
            .get(name)
            .ok_or_else(|| XlexError::SheetNotFound {
                name: name.to_string(),
            })?;

        // Remove from sheets
        self.sheets.remove(index);

        // Rebuild sheet map
        self.sheet_map.clear();
        for (i, sheet) in self.sheets.iter_mut().enumerate() {
            sheet.info.index = i;
            self.sheet_map.insert(sheet.name().to_string(), i);
        }

        // Adjust active sheet
        if self.active_sheet >= self.sheets.len() {
            self.active_sheet = self.sheets.len() - 1;
        }

        self.modified = true;
        Ok(())
    }

    /// Moves a sheet to a new position (0-based index).
    pub fn move_sheet(&mut self, name: &str, new_position: usize) -> XlexResult<()> {
        // Check sheet exists
        let current_index = *self
            .sheet_map
            .get(name)
            .ok_or_else(|| XlexError::SheetNotFound {
                name: name.to_string(),
            })?;

        // Validate new position
        if new_position >= self.sheets.len() {
            return Err(XlexError::SheetIndexOutOfBounds {
                index: new_position,
            });
        }

        // If same position, nothing to do
        if current_index == new_position {
            return Ok(());
        }

        // Remove sheet from current position
        let sheet = self.sheets.remove(current_index);

        // Insert at new position
        self.sheets.insert(new_position, sheet);

        // Rebuild sheet map with updated indices
        self.sheet_map.clear();
        for (i, sheet) in self.sheets.iter_mut().enumerate() {
            sheet.info.index = i;
            self.sheet_map.insert(sheet.name().to_string(), i);
        }

        // Adjust active sheet if needed
        if self.active_sheet == current_index {
            self.active_sheet = new_position;
        } else if current_index < self.active_sheet && new_position >= self.active_sheet {
            self.active_sheet -= 1;
        } else if current_index > self.active_sheet && new_position <= self.active_sheet {
            self.active_sheet += 1;
        }

        self.modified = true;
        Ok(())
    }

    /// Renames a sheet.
    pub fn rename_sheet(&mut self, old_name: &str, new_name: &str) -> XlexResult<()> {
        // Validate new name
        Self::validate_sheet_name(new_name)?;

        // Check old name exists
        let index = *self
            .sheet_map
            .get(old_name)
            .ok_or_else(|| XlexError::SheetNotFound {
                name: old_name.to_string(),
            })?;

        // Check new name doesn't exist (unless same as old)
        if old_name != new_name && self.sheet_map.contains_key(new_name) {
            return Err(XlexError::SheetAlreadyExists {
                name: new_name.to_string(),
            });
        }

        // Update
        self.sheet_map.remove(old_name);
        self.sheet_map.insert(new_name.to_string(), index);
        self.sheets[index].set_name(new_name);
        self.modified = true;

        Ok(())
    }

    /// Validates a sheet name.
    fn validate_sheet_name(name: &str) -> XlexResult<()> {
        if name.is_empty() {
            return Err(XlexError::InvalidSheetName {
                name: name.to_string(),
                reason: "Sheet name cannot be empty".to_string(),
            });
        }

        if name.len() > 31 {
            return Err(XlexError::InvalidSheetName {
                name: name.to_string(),
                reason: "Sheet name cannot exceed 31 characters".to_string(),
            });
        }

        let invalid_chars = [':', '\\', '/', '?', '*', '[', ']'];
        for c in invalid_chars {
            if name.contains(c) {
                return Err(XlexError::InvalidSheetName {
                    name: name.to_string(),
                    reason: format!("Sheet name cannot contain '{}'", c),
                });
            }
        }

        if name.starts_with('\'') || name.ends_with('\'') {
            return Err(XlexError::InvalidSheetName {
                name: name.to_string(),
                reason: "Sheet name cannot start or end with apostrophe".to_string(),
            });
        }

        Ok(())
    }

    /// Gets the active sheet index.
    pub fn active_sheet_index(&self) -> usize {
        self.active_sheet
    }

    /// Sets the active sheet by index.
    pub fn set_active_sheet(&mut self, index: usize) -> XlexResult<()> {
        if index >= self.sheets.len() {
            return Err(XlexError::SheetIndexOutOfBounds { index });
        }
        self.active_sheet = index;
        self.modified = true;
        Ok(())
    }

    /// Sets the active sheet by name.
    pub fn set_active_sheet_by_name(&mut self, name: &str) -> XlexResult<()> {
        let index = *self
            .sheet_map
            .get(name)
            .ok_or_else(|| XlexError::SheetNotFound {
                name: name.to_string(),
            })?;
        self.active_sheet = index;
        self.modified = true;
        Ok(())
    }

    /// Gets the visibility of a sheet.
    pub fn get_sheet_visibility(&self, name: &str) -> XlexResult<SheetVisibility> {
        let sheet = self
            .get_sheet(name)
            .ok_or_else(|| XlexError::SheetNotFound {
                name: name.to_string(),
            })?;
        Ok(sheet.info.visibility)
    }

    /// Sets the visibility of a sheet.
    pub fn set_sheet_visibility(
        &mut self,
        name: &str,
        visibility: SheetVisibility,
    ) -> XlexResult<()> {
        let sheet = self
            .get_sheet_mut(name)
            .ok_or_else(|| XlexError::SheetNotFound {
                name: name.to_string(),
            })?;
        sheet.info.visibility = visibility;
        Ok(())
    }

    /// Gets a cell value.
    pub fn get_cell(&self, sheet_name: &str, cell_ref: &CellRef) -> XlexResult<CellValue> {
        let sheet = self
            .get_sheet(sheet_name)
            .ok_or_else(|| XlexError::SheetNotFound {
                name: sheet_name.to_string(),
            })?;
        Ok(sheet.get_value(cell_ref))
    }

    /// Sets a cell value.
    pub fn set_cell(
        &mut self,
        sheet_name: &str,
        cell_ref: CellRef,
        value: CellValue,
    ) -> XlexResult<()> {
        let sheet = self
            .get_sheet_mut(sheet_name)
            .ok_or_else(|| XlexError::SheetNotFound {
                name: sheet_name.to_string(),
            })?;
        sheet.set_cell(cell_ref, value);
        Ok(())
    }

    /// Clears a cell.
    pub fn clear_cell(&mut self, sheet_name: &str, cell_ref: &CellRef) -> XlexResult<()> {
        let sheet = self
            .get_sheet_mut(sheet_name)
            .ok_or_else(|| XlexError::SheetNotFound {
                name: sheet_name.to_string(),
            })?;
        sheet.clear_cell(cell_ref);
        Ok(())
    }

    /// Returns the style registry.
    pub fn style_registry(&self) -> &StyleRegistry {
        &self.style_registry
    }

    /// Returns a mutable reference to the style registry.
    pub fn style_registry_mut(&mut self) -> &mut StyleRegistry {
        self.modified = true;
        &mut self.style_registry
    }

    /// Returns the shared strings.
    pub fn shared_strings(&self) -> &[String] {
        &self.shared_strings
    }

    /// Adds a shared string and returns its index.
    pub fn add_shared_string(&mut self, s: impl Into<String>) -> usize {
        let s = s.into();
        // TODO: Use a hash map for deduplication
        self.shared_strings.push(s);
        self.modified = true;
        self.shared_strings.len() - 1
    }

    /// Calculates workbook statistics.
    pub fn stats(&self) -> WorkbookStats {
        let mut total_cells = 0;
        let mut formula_count = 0;

        for sheet in &self.sheets {
            total_cells += sheet.cell_count();
            for cell in sheet.cells() {
                if matches!(cell.value, CellValue::Formula { .. }) {
                    formula_count += 1;
                }
            }
        }

        WorkbookStats {
            sheet_count: self.sheets.len(),
            total_cells,
            formula_count,
            style_count: self.style_registry.len(),
            string_count: self.shared_strings.len(),
            file_size: self
                .path
                .as_ref()
                .and_then(|p| p.metadata().ok())
                .map(|m| m.len())
                .unwrap_or(0),
        }
    }

    /// Saves the workbook to its original path.
    pub fn save(&self) -> XlexResult<()> {
        let path = self
            .path
            .as_ref()
            .ok_or_else(|| XlexError::OperationFailed {
                message: "No file path set for workbook".to_string(),
            })?;
        self.save_as(path)
    }

    /// Saves the workbook to a new path.
    pub fn save_as(&self, path: impl AsRef<Path>) -> XlexResult<()> {
        let path = path.as_ref();

        // Check extension
        if path.extension().and_then(|e| e.to_str()) != Some("xlsx") {
            return Err(XlexError::InvalidExtension {
                path: path.to_path_buf(),
            });
        }

        // Create writer
        let writer = crate::writer::WorkbookWriter::new();
        writer.write(self, path)
    }

    /// Returns true if the workbook has been modified.
    pub fn is_modified(&self) -> bool {
        self.modified
    }

    /// Returns all defined names (named ranges).
    pub fn defined_names(&self) -> &[DefinedName] {
        &self.defined_names
    }

    /// Gets a defined name by name.
    pub fn get_defined_name(&self, name: &str) -> Option<&DefinedName> {
        self.defined_names.iter().find(|d| d.name == name)
    }

    /// Adds or updates a defined name.
    pub fn set_defined_name(&mut self, defined_name: DefinedName) {
        self.modified = true;
        // Remove existing with same name and scope
        self.defined_names.retain(|d| {
            !(d.name == defined_name.name && d.local_sheet_id == defined_name.local_sheet_id)
        });
        self.defined_names.push(defined_name);
    }

    /// Removes a defined name.
    pub fn remove_defined_name(&mut self, name: &str) -> bool {
        let before_len = self.defined_names.len();
        self.defined_names.retain(|d| d.name != name);
        let removed = self.defined_names.len() < before_len;
        if removed {
            self.modified = true;
        }
        removed
    }

    /// Internal constructor for the parser.
    /// This is hidden from the public API and should only be used by the parser module.
    #[doc(hidden)]
    #[allow(clippy::too_many_arguments)]
    pub fn __from_parts(
        path: Option<PathBuf>,
        properties: DocumentProperties,
        sheets: Vec<Sheet>,
        sheet_map: HashMap<String, usize>,
        style_registry: StyleRegistry,
        shared_strings: Vec<String>,
        defined_names: Vec<DefinedName>,
        active_sheet: usize,
        modified: bool,
    ) -> Self {
        Self {
            path,
            properties,
            sheets,
            sheet_map,
            style_registry,
            shared_strings,
            defined_names,
            active_sheet,
            modified,
        }
    }
}

impl Default for Workbook {
    fn default() -> Self {
        Self::new()
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_new_workbook() {
        let wb = Workbook::new();
        assert_eq!(wb.sheet_count(), 1);
        assert_eq!(wb.sheet_names(), vec!["Sheet1"]);
        assert!(wb.is_modified());
        assert!(wb.path().is_none());
    }

    #[test]
    fn test_workbook_default() {
        let wb = Workbook::default();
        assert_eq!(wb.sheet_count(), 1);
        assert_eq!(wb.sheet_names(), vec!["Sheet1"]);
    }

    #[test]
    fn test_workbook_with_sheets() {
        let wb = Workbook::with_sheets(&["Data", "Summary", "Config"]);
        assert_eq!(wb.sheet_count(), 3);
        assert_eq!(wb.sheet_names(), vec!["Data", "Summary", "Config"]);
    }

    #[test]
    fn test_workbook_with_empty_sheets() {
        // Should create default sheet when empty array provided
        let wb = Workbook::with_sheets(&[]);
        assert_eq!(wb.sheet_count(), 1);
        assert_eq!(wb.sheet_names(), vec!["Sheet1"]);
    }

    #[test]
    fn test_add_sheet() {
        let mut wb = Workbook::new();
        let index = wb.add_sheet("NewSheet").unwrap();
        assert_eq!(index, 1);
        assert_eq!(wb.sheet_count(), 2);
        assert!(wb.get_sheet("NewSheet").is_some());
    }

    #[test]
    fn test_add_duplicate_sheet() {
        let mut wb = Workbook::new();
        let result = wb.add_sheet("Sheet1");
        assert!(result.is_err());
        assert!(matches!(
            result.unwrap_err(),
            XlexError::SheetAlreadyExists { .. }
        ));
    }

    #[test]
    fn test_remove_sheet() {
        let mut wb = Workbook::with_sheets(&["Sheet1", "Sheet2"]);
        wb.remove_sheet("Sheet1").unwrap();
        assert_eq!(wb.sheet_count(), 1);
        assert!(wb.get_sheet("Sheet1").is_none());
        assert!(wb.get_sheet("Sheet2").is_some());
    }

    #[test]
    fn test_remove_nonexistent_sheet() {
        let mut wb = Workbook::with_sheets(&["Sheet1", "Sheet2"]);
        let result = wb.remove_sheet("NonExistent");
        assert!(result.is_err());
        assert!(matches!(
            result.unwrap_err(),
            XlexError::SheetNotFound { .. }
        ));
    }

    #[test]
    fn test_remove_last_sheet() {
        let mut wb = Workbook::new();
        let result = wb.remove_sheet("Sheet1");
        assert!(result.is_err());
        assert!(matches!(
            result.unwrap_err(),
            XlexError::CannotDeleteLastSheet
        ));
    }

    #[test]
    fn test_rename_sheet() {
        let mut wb = Workbook::new();
        wb.rename_sheet("Sheet1", "Data").unwrap();
        assert!(wb.get_sheet("Sheet1").is_none());
        assert!(wb.get_sheet("Data").is_some());
    }

    #[test]
    fn test_rename_sheet_to_same_name() {
        let mut wb = Workbook::new();
        wb.rename_sheet("Sheet1", "Sheet1").unwrap(); // Should succeed
        assert!(wb.get_sheet("Sheet1").is_some());
    }

    #[test]
    fn test_rename_sheet_to_existing_name() {
        let mut wb = Workbook::with_sheets(&["Sheet1", "Sheet2"]);
        let result = wb.rename_sheet("Sheet1", "Sheet2");
        assert!(result.is_err());
        assert!(matches!(
            result.unwrap_err(),
            XlexError::SheetAlreadyExists { .. }
        ));
    }

    #[test]
    fn test_rename_nonexistent_sheet() {
        let mut wb = Workbook::new();
        let result = wb.rename_sheet("NonExistent", "NewName");
        assert!(result.is_err());
        assert!(matches!(
            result.unwrap_err(),
            XlexError::SheetNotFound { .. }
        ));
    }

    #[test]
    fn test_validate_sheet_name() {
        assert!(Workbook::validate_sheet_name("Valid Name").is_ok());
        assert!(Workbook::validate_sheet_name("Sheet1").is_ok());
        assert!(Workbook::validate_sheet_name("a".repeat(31).as_str()).is_ok()); // Max length

        // Invalid cases
        assert!(Workbook::validate_sheet_name("").is_err());
        assert!(Workbook::validate_sheet_name("a".repeat(32).as_str()).is_err());
        assert!(Workbook::validate_sheet_name("Invalid:Name").is_err());
        assert!(Workbook::validate_sheet_name("Invalid\\Name").is_err());
        assert!(Workbook::validate_sheet_name("Invalid/Name").is_err());
        assert!(Workbook::validate_sheet_name("Invalid?Name").is_err());
        assert!(Workbook::validate_sheet_name("Invalid*Name").is_err());
        assert!(Workbook::validate_sheet_name("Invalid[Name").is_err());
        assert!(Workbook::validate_sheet_name("Invalid]Name").is_err());
        assert!(Workbook::validate_sheet_name("'Name").is_err());
        assert!(Workbook::validate_sheet_name("Name'").is_err());
    }

    #[test]
    fn test_move_sheet() {
        let mut wb = Workbook::with_sheets(&["A", "B", "C", "D"]);

        // Move C (index 2) to index 0
        wb.move_sheet("C", 0).unwrap();
        assert_eq!(wb.sheet_names(), vec!["C", "A", "B", "D"]);

        // Move C (now index 0) to index 3
        wb.move_sheet("C", 3).unwrap();
        assert_eq!(wb.sheet_names(), vec!["A", "B", "D", "C"]);
    }

    #[test]
    fn test_move_sheet_same_position() {
        let mut wb = Workbook::with_sheets(&["A", "B", "C"]);
        wb.move_sheet("B", 1).unwrap(); // Same position
        assert_eq!(wb.sheet_names(), vec!["A", "B", "C"]);
    }

    #[test]
    fn test_move_sheet_invalid_position() {
        let mut wb = Workbook::with_sheets(&["A", "B", "C"]);
        let result = wb.move_sheet("A", 5);
        assert!(result.is_err());
        assert!(matches!(
            result.unwrap_err(),
            XlexError::SheetIndexOutOfBounds { .. }
        ));
    }

    #[test]
    fn test_move_nonexistent_sheet() {
        let mut wb = Workbook::new();
        let result = wb.move_sheet("NonExistent", 0);
        assert!(result.is_err());
        assert!(matches!(
            result.unwrap_err(),
            XlexError::SheetNotFound { .. }
        ));
    }

    #[test]
    fn test_cell_operations() {
        let mut wb = Workbook::new();
        let cell_ref = CellRef::new(1, 1);

        wb.set_cell("Sheet1", cell_ref.clone(), CellValue::string("Hello"))
            .unwrap();

        let value = wb.get_cell("Sheet1", &cell_ref).unwrap();
        assert_eq!(value, CellValue::String("Hello".to_string()));

        wb.clear_cell("Sheet1", &cell_ref).unwrap();
        let value = wb.get_cell("Sheet1", &cell_ref).unwrap();
        assert!(value.is_empty());
    }

    #[test]
    fn test_cell_operations_nonexistent_sheet() {
        let mut wb = Workbook::new();
        let cell_ref = CellRef::new(1, 1);

        let result = wb.get_cell("NonExistent", &cell_ref);
        assert!(result.is_err());

        let result = wb.set_cell("NonExistent", cell_ref.clone(), CellValue::string("test"));
        assert!(result.is_err());

        let result = wb.clear_cell("NonExistent", &cell_ref);
        assert!(result.is_err());
    }

    #[test]
    fn test_workbook_stats() {
        let mut wb = Workbook::new();
        wb.set_cell("Sheet1", CellRef::new(1, 1), CellValue::string("Hello"))
            .unwrap();
        wb.set_cell("Sheet1", CellRef::new(1, 2), CellValue::formula("A1"))
            .unwrap();

        let stats = wb.stats();
        assert_eq!(stats.sheet_count, 1);
        assert_eq!(stats.total_cells, 2);
        assert_eq!(stats.formula_count, 1);
    }

    #[test]
    fn test_active_sheet() {
        let mut wb = Workbook::with_sheets(&["A", "B", "C"]);

        assert_eq!(wb.active_sheet_index(), 0);

        wb.set_active_sheet(2).unwrap();
        assert_eq!(wb.active_sheet_index(), 2);

        wb.set_active_sheet_by_name("A").unwrap();
        assert_eq!(wb.active_sheet_index(), 0);
    }

    #[test]
    fn test_active_sheet_invalid_index() {
        let mut wb = Workbook::new();
        let result = wb.set_active_sheet(5);
        assert!(result.is_err());
        assert!(matches!(
            result.unwrap_err(),
            XlexError::SheetIndexOutOfBounds { .. }
        ));
    }

    #[test]
    fn test_active_sheet_nonexistent_name() {
        let mut wb = Workbook::new();
        let result = wb.set_active_sheet_by_name("NonExistent");
        assert!(result.is_err());
        assert!(matches!(
            result.unwrap_err(),
            XlexError::SheetNotFound { .. }
        ));
    }

    #[test]
    fn test_sheet_visibility() {
        let mut wb = Workbook::new();

        // Default is visible
        assert_eq!(
            wb.get_sheet_visibility("Sheet1").unwrap(),
            SheetVisibility::Visible
        );

        wb.set_sheet_visibility("Sheet1", SheetVisibility::Hidden)
            .unwrap();
        assert_eq!(
            wb.get_sheet_visibility("Sheet1").unwrap(),
            SheetVisibility::Hidden
        );

        wb.set_sheet_visibility("Sheet1", SheetVisibility::VeryHidden)
            .unwrap();
        assert_eq!(
            wb.get_sheet_visibility("Sheet1").unwrap(),
            SheetVisibility::VeryHidden
        );
    }

    #[test]
    fn test_sheet_visibility_nonexistent() {
        let mut wb = Workbook::new();

        let result = wb.get_sheet_visibility("NonExistent");
        assert!(result.is_err());

        let result = wb.set_sheet_visibility("NonExistent", SheetVisibility::Hidden);
        assert!(result.is_err());
    }

    #[test]
    fn test_get_sheet_by_index() {
        let wb = Workbook::with_sheets(&["A", "B", "C"]);

        assert!(wb.get_sheet_by_index(0).is_some());
        assert_eq!(wb.get_sheet_by_index(0).unwrap().name(), "A");
        assert_eq!(wb.get_sheet_by_index(1).unwrap().name(), "B");
        assert_eq!(wb.get_sheet_by_index(2).unwrap().name(), "C");
        assert!(wb.get_sheet_by_index(3).is_none());
    }

    #[test]
    fn test_get_sheet_by_index_mut() {
        let mut wb = Workbook::new();

        if let Some(sheet) = wb.get_sheet_by_index_mut(0) {
            sheet.set_cell(CellRef::new(1, 1), CellValue::string("test"));
        }

        let value = wb.get_cell("Sheet1", &CellRef::new(1, 1)).unwrap();
        assert_eq!(value, CellValue::string("test"));

        assert!(wb.get_sheet_by_index_mut(5).is_none());
    }

    #[test]
    fn test_defined_names() {
        let mut wb = Workbook::new();

        assert!(wb.defined_names().is_empty());

        let name1 = DefinedName::new("MyRange", "Sheet1!$A$1:$B$10");
        wb.set_defined_name(name1);

        assert_eq!(wb.defined_names().len(), 1);
        assert!(wb.get_defined_name("MyRange").is_some());
        assert!(wb.get_defined_name("NonExistent").is_none());

        let name2 = DefinedName::new("MyRange", "Sheet1!$C$1:$D$10"); // Same name, update
        wb.set_defined_name(name2);
        assert_eq!(wb.defined_names().len(), 1);
        assert_eq!(
            wb.get_defined_name("MyRange").unwrap().reference,
            "Sheet1!$C$1:$D$10"
        );

        assert!(wb.remove_defined_name("MyRange"));
        assert!(wb.defined_names().is_empty());
        assert!(!wb.remove_defined_name("NonExistent")); // Returns false for non-existent
    }

    #[test]
    fn test_defined_name_with_sheet_scope() {
        let name = DefinedName::with_sheet_scope("LocalName", "Sheet1!$A$1", 0);
        assert_eq!(name.name, "LocalName");
        assert_eq!(name.reference, "Sheet1!$A$1");
        assert_eq!(name.local_sheet_id, Some(0));
        assert!(!name.hidden);
    }

    #[test]
    fn test_properties() {
        let mut wb = Workbook::new();

        assert!(wb.properties().title.is_none());

        wb.properties_mut().title = Some("My Workbook".to_string());
        wb.properties_mut().creator = Some("Test User".to_string());

        assert_eq!(wb.properties().title, Some("My Workbook".to_string()));
        assert_eq!(wb.properties().creator, Some("Test User".to_string()));
    }

    #[test]
    fn test_shared_strings() {
        let mut wb = Workbook::new();

        assert!(wb.shared_strings().is_empty());

        let idx1 = wb.add_shared_string("Hello");
        assert_eq!(idx1, 0);

        let idx2 = wb.add_shared_string("World");
        assert_eq!(idx2, 1);

        assert_eq!(wb.shared_strings().len(), 2);
        assert_eq!(wb.shared_strings()[0], "Hello");
        assert_eq!(wb.shared_strings()[1], "World");
    }

    #[test]
    fn test_style_registry() {
        let mut wb = Workbook::new();

        assert!(wb.style_registry().is_empty());

        let style = crate::style::Style::default();
        let id = wb.style_registry_mut().add(style);

        assert_eq!(id, 0);
        assert_eq!(wb.style_registry().len(), 1);
    }

    #[test]
    fn test_open_nonexistent_file() {
        let result = Workbook::open("/nonexistent/path/file.xlsx");
        assert!(result.is_err());
        assert!(matches!(
            result.unwrap_err(),
            XlexError::FileNotFound { .. }
        ));
    }

    #[test]
    fn test_open_invalid_extension() {
        // Create a temp file with wrong extension
        let temp_dir = std::env::temp_dir();
        let file_path = temp_dir.join("test_invalid.txt");
        std::fs::write(&file_path, "not xlsx").unwrap();

        let result = Workbook::open(&file_path);
        assert!(result.is_err());
        assert!(matches!(
            result.unwrap_err(),
            XlexError::InvalidExtension { .. }
        ));

        std::fs::remove_file(file_path).ok();
    }

    #[test]
    fn test_remove_sheet_adjusts_active_sheet() {
        let mut wb = Workbook::with_sheets(&["A", "B", "C"]);
        wb.set_active_sheet(2).unwrap(); // C is active

        wb.remove_sheet("C").unwrap();
        // Active sheet should be adjusted
        assert!(wb.active_sheet_index() < wb.sheet_count());
    }

    #[test]
    fn test_save_and_open_roundtrip() {
        let temp_dir = std::env::temp_dir();
        let file_path = temp_dir.join("test_roundtrip.xlsx");

        // Create workbook with content
        {
            let mut wb = Workbook::new();

            // Add cells
            let cell_ref = crate::cell::CellRef::parse("A1").unwrap();
            wb.set_cell("Sheet1", cell_ref, CellValue::String("Hello".to_string()))
                .unwrap();

            let cell_ref = crate::cell::CellRef::parse("B1").unwrap();
            wb.set_cell("Sheet1", cell_ref, CellValue::Number(42.5))
                .unwrap();

            let cell_ref = crate::cell::CellRef::parse("C1").unwrap();
            wb.set_cell("Sheet1", cell_ref, CellValue::Boolean(true))
                .unwrap();

            // Save
            wb.save_as(&file_path).unwrap();
        }

        // Reopen and verify
        {
            let wb = Workbook::open(&file_path).unwrap();

            assert_eq!(wb.sheet_count(), 1);
            assert_eq!(wb.sheet_names()[0], "Sheet1");

            // Verify cells - note inline string behavior
            let cell_ref = crate::cell::CellRef::parse("B1").unwrap();
            let value = wb.get_cell("Sheet1", &cell_ref).unwrap();
            assert_eq!(value, CellValue::Number(42.5));

            let cell_ref = crate::cell::CellRef::parse("C1").unwrap();
            let value = wb.get_cell("Sheet1", &cell_ref).unwrap();
            assert_eq!(value, CellValue::Boolean(true));
        }

        // Cleanup
        std::fs::remove_file(file_path).ok();
    }

    #[test]
    fn test_save_with_multiple_sheets() {
        let temp_dir = std::env::temp_dir();
        let file_path = temp_dir.join("test_multi_sheet.xlsx");

        {
            let mut wb = Workbook::with_sheets(&["Data", "Summary", "Config"]);

            // Add content to each sheet
            let cell_ref = crate::cell::CellRef::parse("A1").unwrap();
            wb.set_cell(
                "Data",
                cell_ref,
                CellValue::String("Data Sheet".to_string()),
            )
            .unwrap();

            let cell_ref = crate::cell::CellRef::parse("A1").unwrap();
            wb.set_cell("Summary", cell_ref, CellValue::Number(100.0))
                .unwrap();

            wb.save_as(&file_path).unwrap();
        }

        {
            let wb = Workbook::open(&file_path).unwrap();

            assert_eq!(wb.sheet_count(), 3);
            let names = wb.sheet_names();
            assert!(names.contains(&"Data"));
            assert!(names.contains(&"Summary"));
            assert!(names.contains(&"Config"));
        }

        std::fs::remove_file(file_path).ok();
    }

    #[test]
    fn test_save_with_properties() {
        let temp_dir = std::env::temp_dir();
        let file_path = temp_dir.join("test_properties.xlsx");

        {
            let mut wb = Workbook::new();

            wb.properties_mut().title = Some("Test Title".to_string());
            wb.properties_mut().creator = Some("Test Author".to_string());
            wb.properties_mut().subject = Some("Test Subject".to_string());

            wb.save_as(&file_path).unwrap();
        }

        {
            let wb = Workbook::open(&file_path).unwrap();

            assert_eq!(wb.properties().title, Some("Test Title".to_string()));
            assert_eq!(wb.properties().creator, Some("Test Author".to_string()));
            assert_eq!(wb.properties().subject, Some("Test Subject".to_string()));
        }

        std::fs::remove_file(file_path).ok();
    }

    #[test]
    fn test_save_as_invalid_extension() {
        let temp_dir = std::env::temp_dir();
        let file_path = temp_dir.join("test_invalid.xls");

        let wb = Workbook::new();
        let result = wb.save_as(&file_path);

        assert!(result.is_err());
        assert!(matches!(
            result.unwrap_err(),
            XlexError::InvalidExtension { .. }
        ));
    }

    #[test]
    fn test_save_no_path() {
        let wb = Workbook::new();
        let result = wb.save();

        assert!(result.is_err());
        assert!(matches!(
            result.unwrap_err(),
            XlexError::OperationFailed { .. }
        ));
    }

    #[test]
    fn test_from_reader() {
        let temp_dir = std::env::temp_dir();
        let file_path = temp_dir.join("test_from_reader.xlsx");

        // Create a file first
        {
            let mut wb = Workbook::new();
            let cell_ref = crate::cell::CellRef::parse("A1").unwrap();
            wb.set_cell("Sheet1", cell_ref, CellValue::Number(123.0))
                .unwrap();
            wb.save_as(&file_path).unwrap();
        }

        // Read using from_reader
        {
            let file = std::fs::File::open(&file_path).unwrap();
            let reader = std::io::BufReader::new(file);
            let wb = Workbook::from_reader(reader, None).unwrap();

            assert_eq!(wb.sheet_count(), 1);
            assert!(wb.path().is_none());

            let cell_ref = crate::cell::CellRef::parse("A1").unwrap();
            let value = wb.get_cell("Sheet1", &cell_ref).unwrap();
            assert_eq!(value, CellValue::Number(123.0));
        }

        std::fs::remove_file(file_path).ok();
    }

    #[test]
    fn test_clear_cell() {
        let mut wb = Workbook::new();

        let cell_ref = crate::cell::CellRef::parse("A1").unwrap();
        wb.set_cell("Sheet1", cell_ref.clone(), CellValue::Number(42.0))
            .unwrap();

        assert_eq!(
            wb.get_cell("Sheet1", &cell_ref).unwrap(),
            CellValue::Number(42.0)
        );

        wb.clear_cell("Sheet1", &cell_ref).unwrap();

        assert_eq!(wb.get_cell("Sheet1", &cell_ref).unwrap(), CellValue::Empty);
    }

    #[test]
    fn test_clear_cell_nonexistent_sheet() {
        let mut wb = Workbook::new();
        let cell_ref = crate::cell::CellRef::parse("A1").unwrap();

        let result = wb.clear_cell("NonExistent", &cell_ref);
        assert!(result.is_err());
        assert!(matches!(
            result.unwrap_err(),
            XlexError::SheetNotFound { .. }
        ));
    }

    #[test]
    fn test_hidden_sheet_roundtrip() {
        let temp_dir = std::env::temp_dir();
        let file_path = temp_dir.join("test_hidden_sheet.xlsx");

        {
            let mut wb = Workbook::with_sheets(&["Visible", "Hidden"]);
            wb.set_sheet_visibility("Hidden", crate::sheet::SheetVisibility::Hidden)
                .unwrap();
            wb.save_as(&file_path).unwrap();
        }

        {
            let wb = Workbook::open(&file_path).unwrap();

            assert_eq!(
                wb.get_sheet_visibility("Visible").unwrap(),
                crate::sheet::SheetVisibility::Visible
            );
            assert_eq!(
                wb.get_sheet_visibility("Hidden").unwrap(),
                crate::sheet::SheetVisibility::Hidden
            );
        }

        std::fs::remove_file(file_path).ok();
    }
}
