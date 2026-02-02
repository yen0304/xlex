//! Workbook type and operations.

use std::collections::HashMap;
use std::fs::File;
use std::io::{BufReader, Read, Seek};
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
    pub fn open(path: impl AsRef<Path>) -> XlexResult<Self> {
        let path = path.as_ref();

        // Check file exists
        if !path.exists() {
            return Err(XlexError::FileNotFound {
                path: path.to_path_buf(),
            });
        }

        // Check extension
        if path.extension().and_then(|e| e.to_str()) != Some("xlsx") {
            return Err(XlexError::InvalidExtension {
                path: path.to_path_buf(),
            });
        }

        // Open file
        let file = File::open(path).map_err(|e| {
            if e.kind() == std::io::ErrorKind::PermissionDenied {
                XlexError::PermissionDenied {
                    path: path.to_path_buf(),
                }
            } else {
                XlexError::from(e)
            }
        })?;

        let reader = BufReader::new(file);
        Self::from_reader(reader, Some(path.to_path_buf()))
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

        let index = *self.sheet_map.get(name).ok_or_else(|| XlexError::SheetNotFound {
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
        let current_index = *self.sheet_map.get(name).ok_or_else(|| XlexError::SheetNotFound {
            name: name.to_string(),
        })?;

        // Validate new position
        if new_position >= self.sheets.len() {
            return Err(XlexError::SheetIndexOutOfBounds { index: new_position });
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
        let index = *self.sheet_map.get(name).ok_or_else(|| XlexError::SheetNotFound {
            name: name.to_string(),
        })?;
        self.active_sheet = index;
        self.modified = true;
        Ok(())
    }

    /// Gets the visibility of a sheet.
    pub fn get_sheet_visibility(&self, name: &str) -> XlexResult<SheetVisibility> {
        let sheet = self.get_sheet(name).ok_or_else(|| XlexError::SheetNotFound {
            name: name.to_string(),
        })?;
        Ok(sheet.info.visibility)
    }

    /// Sets the visibility of a sheet.
    pub fn set_sheet_visibility(&mut self, name: &str, visibility: SheetVisibility) -> XlexResult<()> {
        let sheet = self.get_sheet_mut(name).ok_or_else(|| XlexError::SheetNotFound {
            name: name.to_string(),
        })?;
        sheet.info.visibility = visibility;
        Ok(())
    }

    /// Gets a cell value.
    pub fn get_cell(&self, sheet_name: &str, cell_ref: &CellRef) -> XlexResult<CellValue> {
        let sheet = self.get_sheet(sheet_name).ok_or_else(|| XlexError::SheetNotFound {
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
        let path = self.path.as_ref().ok_or_else(|| XlexError::OperationFailed {
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
    }

    #[test]
    fn test_workbook_with_sheets() {
        let wb = Workbook::with_sheets(&["Data", "Summary", "Config"]);
        assert_eq!(wb.sheet_count(), 3);
        assert_eq!(wb.sheet_names(), vec!["Data", "Summary", "Config"]);
    }

    #[test]
    fn test_add_sheet() {
        let mut wb = Workbook::new();
        wb.add_sheet("NewSheet").unwrap();
        assert_eq!(wb.sheet_count(), 2);
        assert!(wb.get_sheet("NewSheet").is_some());
    }

    #[test]
    fn test_add_duplicate_sheet() {
        let mut wb = Workbook::new();
        assert!(wb.add_sheet("Sheet1").is_err());
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
    fn test_remove_last_sheet() {
        let mut wb = Workbook::new();
        assert!(wb.remove_sheet("Sheet1").is_err());
    }

    #[test]
    fn test_rename_sheet() {
        let mut wb = Workbook::new();
        wb.rename_sheet("Sheet1", "Data").unwrap();
        assert!(wb.get_sheet("Sheet1").is_none());
        assert!(wb.get_sheet("Data").is_some());
    }

    #[test]
    fn test_validate_sheet_name() {
        assert!(Workbook::validate_sheet_name("Valid Name").is_ok());
        assert!(Workbook::validate_sheet_name("").is_err());
        assert!(Workbook::validate_sheet_name("a".repeat(32).as_str()).is_err());
        assert!(Workbook::validate_sheet_name("Invalid:Name").is_err());
        assert!(Workbook::validate_sheet_name("'Name").is_err());
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
    fn test_workbook_stats() {
        let mut wb = Workbook::new();
        wb.set_cell(
            "Sheet1",
            CellRef::new(1, 1),
            CellValue::string("Hello"),
        )
        .unwrap();
        wb.set_cell(
            "Sheet1",
            CellRef::new(1, 2),
            CellValue::formula("A1"),
        )
        .unwrap();

        let stats = wb.stats();
        assert_eq!(stats.sheet_count, 1);
        assert_eq!(stats.total_cells, 2);
        assert_eq!(stats.formula_count, 1);
    }
}
