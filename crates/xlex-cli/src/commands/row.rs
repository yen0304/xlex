//! Row operations.

use anyhow::Result;
use clap::{Parser, Subcommand};
use colored::Colorize;

use xlex_core::Workbook;

use super::{GlobalOptions, OutputFormat};

/// Arguments for row operations.
#[derive(Parser)]
pub struct RowArgs {
    #[command(subcommand)]
    pub command: RowCommand,
}

#[derive(Subcommand)]
pub enum RowCommand {
    /// Get row data
    Get {
        /// Path to the xlsx file
        file: std::path::PathBuf,
        /// Sheet name
        sheet: String,
        /// Row number (1-indexed)
        row: u32,
    },
    /// Append a row
    Append {
        /// Path to the xlsx file
        file: std::path::PathBuf,
        /// Sheet name
        sheet: String,
        /// Values (comma-separated)
        values: String,
    },
    /// Insert a row
    Insert {
        /// Path to the xlsx file
        file: std::path::PathBuf,
        /// Sheet name
        sheet: String,
        /// Row number to insert at (1-indexed)
        row: u32,
    },
    /// Delete a row
    Delete {
        /// Path to the xlsx file
        file: std::path::PathBuf,
        /// Sheet name
        sheet: String,
        /// Row number (1-indexed)
        row: u32,
    },
    /// Copy a row
    Copy {
        /// Path to the xlsx file
        file: std::path::PathBuf,
        /// Sheet name
        sheet: String,
        /// Source row number (1-indexed)
        source: u32,
        /// Destination row number (1-indexed)
        dest: u32,
    },
    /// Move a row
    Move {
        /// Path to the xlsx file
        file: std::path::PathBuf,
        /// Sheet name
        sheet: String,
        /// Source row number (1-indexed)
        source: u32,
        /// Destination row number (1-indexed)
        dest: u32,
    },
    /// Set row height
    Height {
        /// Path to the xlsx file
        file: std::path::PathBuf,
        /// Sheet name
        sheet: String,
        /// Row number (1-indexed)
        row: u32,
        /// Height in points (omit to show current)
        height: Option<f64>,
    },
    /// Hide a row
    Hide {
        /// Path to the xlsx file
        file: std::path::PathBuf,
        /// Sheet name
        sheet: String,
        /// Row number (1-indexed)
        row: u32,
    },
    /// Unhide a row
    Unhide {
        /// Path to the xlsx file
        file: std::path::PathBuf,
        /// Sheet name
        sheet: String,
        /// Row number (1-indexed)
        row: u32,
    },
    /// Find rows matching criteria
    Find {
        /// Path to the xlsx file
        file: std::path::PathBuf,
        /// Sheet name
        sheet: String,
        /// Search pattern
        pattern: String,
        /// Column to search in
        #[arg(long, short = 'c')]
        column: Option<String>,
    },
}

/// Run row operations.
pub fn run(args: &RowArgs, global: &GlobalOptions) -> Result<()> {
    match &args.command {
        RowCommand::Get { file, sheet, row } => get(file, sheet, *row, global),
        RowCommand::Append {
            file,
            sheet,
            values,
        } => append(file, sheet, values, global),
        RowCommand::Insert { file, sheet, row } => insert(file, sheet, *row, global),
        RowCommand::Delete { file, sheet, row } => delete(file, sheet, *row, global),
        RowCommand::Copy {
            file,
            sheet,
            source,
            dest,
        } => copy(file, sheet, *source, *dest, global),
        RowCommand::Move {
            file,
            sheet,
            source,
            dest,
        } => move_row(file, sheet, *source, *dest, global),
        RowCommand::Height {
            file,
            sheet,
            row,
            height: row_height,
        } => height(file, sheet, *row, *row_height, global),
        RowCommand::Hide { file, sheet, row } => hide(file, sheet, *row, global),
        RowCommand::Unhide { file, sheet, row } => unhide(file, sheet, *row, global),
        RowCommand::Find {
            file,
            sheet,
            pattern,
            column,
        } => find(file, sheet, pattern, column.as_deref(), global),
    }
}

fn get(file: &std::path::Path, sheet: &str, row: u32, global: &GlobalOptions) -> Result<()> {
    let workbook = Workbook::open(file)?;
    let sheet_obj =
        workbook
            .get_sheet(sheet)
            .ok_or_else(|| xlex_core::XlexError::SheetNotFound {
                name: sheet.to_string(),
            })?;

    // Collect cells in this row
    let mut row_cells: Vec<_> = sheet_obj
        .cells()
        .filter(|c| c.reference.row == row)
        .collect();
    row_cells.sort_by_key(|c| c.reference.col);

    if global.format == OutputFormat::Json {
        let values: Vec<_> = row_cells
            .iter()
            .map(|c| {
                serde_json::json!({
                    "col": c.reference.col,
                    "ref": c.reference.to_a1(),
                    "value": c.value.to_display_string(),
                    "type": c.value.type_name(),
                })
            })
            .collect();
        println!("{}", serde_json::to_string_pretty(&values)?);
    } else if global.format == OutputFormat::Csv {
        let values: Vec<_> = row_cells
            .iter()
            .map(|c| c.value.to_display_string())
            .collect();
        println!("{}", values.join(","));
    } else {
        for cell in row_cells {
            println!("{}: {}", cell.reference.to_a1().cyan(), cell.value);
        }
    }

    Ok(())
}

fn append(file: &std::path::Path, sheet: &str, values: &str, global: &GlobalOptions) -> Result<()> {
    if global.dry_run {
        println!("Would append row with values: {}", values);
        return Ok(());
    }

    let mut workbook = Workbook::open(file)?;
    let sheet_obj =
        workbook
            .get_sheet_mut(sheet)
            .ok_or_else(|| xlex_core::XlexError::SheetNotFound {
                name: sheet.to_string(),
            })?;

    // Find the next empty row
    let mut max_row: u32 = 0;
    for cell in sheet_obj.cells() {
        if cell.reference.row > max_row {
            max_row = cell.reference.row;
        }
    }
    let new_row = max_row + 1;

    // Parse and set values
    let cell_values: Vec<&str> = values.split(',').collect();
    for (col, val) in cell_values.iter().enumerate() {
        let cell_ref = xlex_core::CellRef::new((col + 1) as u32, new_row);
        let value = super::cell::parse_auto_value(val.trim());
        sheet_obj.set_cell(cell_ref, value);
    }

    let _ = sheet_obj;
    workbook.save()?;

    if !global.quiet {
        println!("Appended row {}", new_row.to_string().green());
    }

    Ok(())
}

fn insert(file: &std::path::Path, sheet: &str, row: u32, global: &GlobalOptions) -> Result<()> {
    if global.dry_run {
        println!("Would insert row at position {}", row);
        return Ok(());
    }

    let mut workbook = Workbook::open(file)?;
    let sheet_obj =
        workbook
            .get_sheet_mut(sheet)
            .ok_or_else(|| xlex_core::XlexError::SheetNotFound {
                name: sheet.to_string(),
            })?;

    sheet_obj.insert_rows(row, 1);
    let _ = sheet_obj;
    workbook.save()?;

    if !global.quiet {
        println!("Inserted row at position {}", row.to_string().green());
    }

    Ok(())
}

fn delete(file: &std::path::Path, sheet: &str, row: u32, global: &GlobalOptions) -> Result<()> {
    if global.dry_run {
        println!("Would delete row {}", row);
        return Ok(());
    }

    let mut workbook = Workbook::open(file)?;
    let sheet_obj =
        workbook
            .get_sheet_mut(sheet)
            .ok_or_else(|| xlex_core::XlexError::SheetNotFound {
                name: sheet.to_string(),
            })?;

    sheet_obj.delete_rows(row, 1);
    let _ = sheet_obj;
    workbook.save()?;

    if !global.quiet {
        println!("Deleted row {}", row.to_string().green());
    }

    Ok(())
}

fn copy(
    file: &std::path::Path,
    sheet: &str,
    source: u32,
    dest: u32,
    global: &GlobalOptions,
) -> Result<()> {
    if global.dry_run {
        println!("Would copy row {} to row {}", source, dest);
        return Ok(());
    }

    let mut workbook = Workbook::open(file)?;

    // Get source row data
    let sheet_obj =
        workbook
            .get_sheet(sheet)
            .ok_or_else(|| xlex_core::XlexError::SheetNotFound {
                name: sheet.to_string(),
            })?;

    // Collect all cells in the source row
    let source_cells: Vec<_> = sheet_obj
        .cells()
        .filter(|c| c.reference.row == source)
        .map(|c| (c.reference.col, c.value.clone()))
        .collect();

    let source_height = sheet_obj.get_row_height(source);
    let _ = sheet_obj;

    // Insert a new row at destination
    let sheet_obj =
        workbook
            .get_sheet_mut(sheet)
            .ok_or_else(|| xlex_core::XlexError::SheetNotFound {
                name: sheet.to_string(),
            })?;

    sheet_obj.insert_rows(dest, 1);

    // Copy cells to destination
    for (col, value) in source_cells {
        let dest_ref = xlex_core::CellRef::new(col, dest);
        sheet_obj.set_cell(dest_ref, value);
    }

    // Copy row height if set
    if let Some(h) = source_height {
        sheet_obj.set_row_height(dest, h);
    }

    let _ = sheet_obj;
    workbook.save()?;

    if !global.quiet {
        println!(
            "Copied row {} to row {}",
            source.to_string().cyan(),
            dest.to_string().green()
        );
    }

    Ok(())
}

fn move_row(
    file: &std::path::Path,
    sheet: &str,
    source: u32,
    dest: u32,
    global: &GlobalOptions,
) -> Result<()> {
    if global.dry_run {
        println!("Would move row {} to row {}", source, dest);
        return Ok(());
    }

    let mut workbook = Workbook::open(file)?;

    // Get source row data
    let sheet_obj =
        workbook
            .get_sheet(sheet)
            .ok_or_else(|| xlex_core::XlexError::SheetNotFound {
                name: sheet.to_string(),
            })?;

    // Collect all cells in the source row
    let source_cells: Vec<_> = sheet_obj
        .cells()
        .filter(|c| c.reference.row == source)
        .map(|c| (c.reference.col, c.value.clone()))
        .collect();

    let source_height = sheet_obj.get_row_height(source);
    let _ = sheet_obj;

    let sheet_obj =
        workbook
            .get_sheet_mut(sheet)
            .ok_or_else(|| xlex_core::XlexError::SheetNotFound {
                name: sheet.to_string(),
            })?;

    // The logic differs based on whether moving up or down
    if source < dest {
        // Moving down: insert first, then delete (adjusted source)
        sheet_obj.insert_rows(dest + 1, 1);
        let actual_dest = dest + 1;

        // Copy cells
        for (col, value) in &source_cells {
            let dest_ref = xlex_core::CellRef::new(*col, actual_dest);
            sheet_obj.set_cell(dest_ref, value.clone());
        }

        if let Some(h) = source_height {
            sheet_obj.set_row_height(actual_dest, h);
        }

        // Delete original row
        sheet_obj.delete_rows(source, 1);
    } else {
        // Moving up: delete first, then insert at adjusted position
        sheet_obj.delete_rows(source, 1);
        sheet_obj.insert_rows(dest, 1);

        // Copy cells
        for (col, value) in &source_cells {
            let dest_ref = xlex_core::CellRef::new(*col, dest);
            sheet_obj.set_cell(dest_ref, value.clone());
        }

        if let Some(h) = source_height {
            sheet_obj.set_row_height(dest, h);
        }
    }

    let _ = sheet_obj;
    workbook.save()?;

    if !global.quiet {
        println!(
            "Moved row {} to row {}",
            source.to_string().cyan(),
            dest.to_string().green()
        );
    }

    Ok(())
}

fn height(
    file: &std::path::Path,
    sheet: &str,
    row: u32,
    height: Option<f64>,
    global: &GlobalOptions,
) -> Result<()> {
    if let Some(h) = height {
        if global.dry_run {
            println!("Would set row {} height to {}", row, h);
            return Ok(());
        }

        let mut workbook = Workbook::open(file)?;
        let sheet_obj =
            workbook
                .get_sheet_mut(sheet)
                .ok_or_else(|| xlex_core::XlexError::SheetNotFound {
                    name: sheet.to_string(),
                })?;

        sheet_obj.set_row_height(row, h);
        let _ = sheet_obj;
        workbook.save()?;

        if !global.quiet {
            println!("Set row {} height to {}", row.to_string().cyan(), h);
        }
    } else {
        let workbook = Workbook::open(file)?;
        let sheet_obj =
            workbook
                .get_sheet(sheet)
                .ok_or_else(|| xlex_core::XlexError::SheetNotFound {
                    name: sheet.to_string(),
                })?;

        if let Some(h) = sheet_obj.get_row_height(row) {
            println!("{}", h);
        } else {
            println!("default");
        }
    }

    Ok(())
}

fn hide(file: &std::path::Path, sheet: &str, row: u32, global: &GlobalOptions) -> Result<()> {
    if global.dry_run {
        println!("Would hide row {}", row);
        return Ok(());
    }

    let mut workbook = Workbook::open(file)?;
    let sheet_obj =
        workbook
            .get_sheet_mut(sheet)
            .ok_or_else(|| xlex_core::XlexError::SheetNotFound {
                name: sheet.to_string(),
            })?;

    sheet_obj.set_row_hidden(row, true);
    let _ = sheet_obj;
    workbook.save()?;

    if !global.quiet {
        println!("Hid row {}", row.to_string().dimmed());
    }

    Ok(())
}

fn unhide(file: &std::path::Path, sheet: &str, row: u32, global: &GlobalOptions) -> Result<()> {
    if global.dry_run {
        println!("Would unhide row {}", row);
        return Ok(());
    }

    let mut workbook = Workbook::open(file)?;
    let sheet_obj =
        workbook
            .get_sheet_mut(sheet)
            .ok_or_else(|| xlex_core::XlexError::SheetNotFound {
                name: sheet.to_string(),
            })?;

    sheet_obj.set_row_hidden(row, false);
    let _ = sheet_obj;
    workbook.save()?;

    if !global.quiet {
        println!("Unhid row {}", row.to_string().green());
    }

    Ok(())
}

fn find(
    file: &std::path::Path,
    sheet: &str,
    pattern: &str,
    column: Option<&str>,
    global: &GlobalOptions,
) -> Result<()> {
    let workbook = Workbook::open(file)?;
    let sheet_obj =
        workbook
            .get_sheet(sheet)
            .ok_or_else(|| xlex_core::XlexError::SheetNotFound {
                name: sheet.to_string(),
            })?;

    let col_filter: Option<u32> =
        column.and_then(|c| xlex_core::CellRef::col_from_letters_pub(&c.to_uppercase()));

    let mut matches: Vec<u32> = Vec::new();
    for cell in sheet_obj.cells() {
        // Apply column filter if specified
        if let Some(col) = col_filter {
            if cell.reference.col != col {
                continue;
            }
        }

        let value = cell.value.to_display_string();
        if value.contains(pattern) {
            if !matches.contains(&cell.reference.row) {
                matches.push(cell.reference.row);
            }
        }
    }

    matches.sort();

    if global.format == OutputFormat::Json {
        let json = serde_json::json!({
            "pattern": pattern,
            "matches": matches,
            "count": matches.len(),
        });
        println!("{}", serde_json::to_string_pretty(&json)?);
    } else {
        for row in matches {
            println!("{}", row);
        }
    }

    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;
    use tempfile::TempDir;
    use xlex_core::CellValue;

    fn default_global() -> GlobalOptions {
        GlobalOptions {
            quiet: true,
            verbose: false,
            format: OutputFormat::Text,
            no_color: true,
            color: false,
            json_errors: false,
            dry_run: false,
            output: None,
        }
    }

    fn create_test_workbook(dir: &TempDir, name: &str) -> std::path::PathBuf {
        let file_path = dir.path().join(name);
        let wb = Workbook::new();
        wb.save_as(&file_path).unwrap();
        file_path
    }

    fn setup_test_data(file: &std::path::Path) {
        let mut wb = Workbook::open(file).unwrap();
        // Set up test data in rows 1-3, columns A-C
        for row in 1..=3 {
            for col in 1..=3 {
                let cell_ref = xlex_core::CellRef::new(col, row);
                let value = CellValue::Number((row * 10 + col) as f64);
                wb.set_cell("Sheet1", cell_ref, value).unwrap();
            }
        }
        wb.save().unwrap();
    }

    #[test]
    fn test_get_row() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "get.xlsx");
        setup_test_data(&file_path);

        let result = get(&file_path, "Sheet1", 1, &default_global());
        assert!(result.is_ok());
    }

    #[test]
    fn test_get_row_json() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "get_json.xlsx");
        setup_test_data(&file_path);

        let mut global = default_global();
        global.format = OutputFormat::Json;

        let result = get(&file_path, "Sheet1", 1, &global);
        assert!(result.is_ok());
    }

    #[test]
    fn test_get_row_csv() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "get_csv.xlsx");
        setup_test_data(&file_path);

        let mut global = default_global();
        global.format = OutputFormat::Csv;

        let result = get(&file_path, "Sheet1", 1, &global);
        assert!(result.is_ok());
    }

    #[test]
    fn test_append_row() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "append.xlsx");

        let result = append(&file_path, "Sheet1", "1,2,3", &default_global());
        assert!(result.is_ok());

        let wb = Workbook::open(&file_path).unwrap();
        let cell_ref = xlex_core::CellRef::new(1, 1);
        let value = wb.get_cell("Sheet1", &cell_ref).unwrap();
        assert_eq!(value, CellValue::Number(1.0));
    }

    #[test]
    fn test_insert_row() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "insert.xlsx");
        setup_test_data(&file_path);

        let result = insert(&file_path, "Sheet1", 2, &default_global());
        assert!(result.is_ok());
    }

    #[test]
    fn test_delete_row() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "delete.xlsx");
        setup_test_data(&file_path);

        let result = delete(&file_path, "Sheet1", 1, &default_global());
        assert!(result.is_ok());
    }

    #[test]
    fn test_copy_row() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "copy.xlsx");
        setup_test_data(&file_path);

        let result = copy(&file_path, "Sheet1", 1, 5, &default_global());
        assert!(result.is_ok());
    }

    #[test]
    fn test_move_row_down() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "move_down.xlsx");
        setup_test_data(&file_path);

        let result = move_row(&file_path, "Sheet1", 1, 5, &default_global());
        assert!(result.is_ok());
    }

    #[test]
    fn test_move_row_up() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "move_up.xlsx");
        setup_test_data(&file_path);

        let result = move_row(&file_path, "Sheet1", 3, 1, &default_global());
        assert!(result.is_ok());
    }

    #[test]
    fn test_height_set() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "height.xlsx");

        let result = height(&file_path, "Sheet1", 1, Some(30.0), &default_global());
        assert!(result.is_ok());
    }

    #[test]
    fn test_height_get() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "height_get.xlsx");

        // Set height first
        height(&file_path, "Sheet1", 1, Some(30.0), &default_global()).unwrap();

        // Then get it
        let result = height(&file_path, "Sheet1", 1, None, &default_global());
        assert!(result.is_ok());
    }

    #[test]
    fn test_hide_row() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "hide.xlsx");

        let result = hide(&file_path, "Sheet1", 1, &default_global());
        assert!(result.is_ok());
    }

    #[test]
    fn test_unhide_row() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "unhide.xlsx");

        // Hide first
        hide(&file_path, "Sheet1", 1, &default_global()).unwrap();

        let result = unhide(&file_path, "Sheet1", 1, &default_global());
        assert!(result.is_ok());
    }

    #[test]
    fn test_find_rows() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "find.xlsx");
        setup_test_data(&file_path);

        let result = find(&file_path, "Sheet1", "1", None, &default_global());
        assert!(result.is_ok());
    }

    #[test]
    fn test_find_rows_with_column() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "find_col.xlsx");
        setup_test_data(&file_path);

        let result = find(&file_path, "Sheet1", "1", Some("A"), &default_global());
        assert!(result.is_ok());
    }

    #[test]
    fn test_dry_run_operations() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "dry.xlsx");

        let mut global = default_global();
        global.dry_run = true;

        assert!(append(&file_path, "Sheet1", "1,2,3", &global).is_ok());
        assert!(insert(&file_path, "Sheet1", 1, &global).is_ok());
        assert!(delete(&file_path, "Sheet1", 1, &global).is_ok());
        assert!(copy(&file_path, "Sheet1", 1, 5, &global).is_ok());
        assert!(move_row(&file_path, "Sheet1", 1, 5, &global).is_ok());
        assert!(height(&file_path, "Sheet1", 1, Some(30.0), &global).is_ok());
        assert!(hide(&file_path, "Sheet1", 1, &global).is_ok());
        assert!(unhide(&file_path, "Sheet1", 1, &global).is_ok());
    }

    // Additional tests for better coverage

    #[test]
    fn test_run_get_command() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "run_get.xlsx");
        setup_test_data(&file_path);

        let args = RowArgs {
            command: RowCommand::Get {
                file: file_path,
                sheet: "Sheet1".to_string(),
                row: 1,
            },
        };

        let result = run(&args, &default_global());
        assert!(result.is_ok());
    }

    #[test]
    fn test_run_append_command() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "run_append.xlsx");

        let args = RowArgs {
            command: RowCommand::Append {
                file: file_path,
                sheet: "Sheet1".to_string(),
                values: "1,2,3".to_string(),
            },
        };

        let result = run(&args, &default_global());
        assert!(result.is_ok());
    }

    #[test]
    fn test_run_insert_command() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "run_insert.xlsx");
        setup_test_data(&file_path);

        let args = RowArgs {
            command: RowCommand::Insert {
                file: file_path,
                sheet: "Sheet1".to_string(),
                row: 2,
            },
        };

        let result = run(&args, &default_global());
        assert!(result.is_ok());
    }

    #[test]
    fn test_run_delete_command() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "run_delete.xlsx");
        setup_test_data(&file_path);

        let args = RowArgs {
            command: RowCommand::Delete {
                file: file_path,
                sheet: "Sheet1".to_string(),
                row: 1,
            },
        };

        let result = run(&args, &default_global());
        assert!(result.is_ok());
    }

    #[test]
    fn test_run_copy_command() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "run_copy.xlsx");
        setup_test_data(&file_path);

        let args = RowArgs {
            command: RowCommand::Copy {
                file: file_path,
                sheet: "Sheet1".to_string(),
                source: 1,
                dest: 5,
            },
        };

        let result = run(&args, &default_global());
        assert!(result.is_ok());
    }

    #[test]
    fn test_run_move_command() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "run_move.xlsx");
        setup_test_data(&file_path);

        let args = RowArgs {
            command: RowCommand::Move {
                file: file_path,
                sheet: "Sheet1".to_string(),
                source: 1,
                dest: 5,
            },
        };

        let result = run(&args, &default_global());
        assert!(result.is_ok());
    }

    #[test]
    fn test_run_height_command() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "run_height.xlsx");

        let args = RowArgs {
            command: RowCommand::Height {
                file: file_path,
                sheet: "Sheet1".to_string(),
                row: 1,
                height: Some(30.0),
            },
        };

        let result = run(&args, &default_global());
        assert!(result.is_ok());
    }

    #[test]
    fn test_run_hide_command() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "run_hide.xlsx");

        let args = RowArgs {
            command: RowCommand::Hide {
                file: file_path,
                sheet: "Sheet1".to_string(),
                row: 1,
            },
        };

        let result = run(&args, &default_global());
        assert!(result.is_ok());
    }

    #[test]
    fn test_run_unhide_command() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "run_unhide.xlsx");

        hide(&file_path, "Sheet1", 1, &default_global()).unwrap();

        let args = RowArgs {
            command: RowCommand::Unhide {
                file: file_path,
                sheet: "Sheet1".to_string(),
                row: 1,
            },
        };

        let result = run(&args, &default_global());
        assert!(result.is_ok());
    }

    #[test]
    fn test_run_find_command() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "run_find.xlsx");
        setup_test_data(&file_path);

        let args = RowArgs {
            command: RowCommand::Find {
                file: file_path,
                sheet: "Sheet1".to_string(),
                pattern: "1".to_string(),
                column: None,
            },
        };

        let result = run(&args, &default_global());
        assert!(result.is_ok());
    }

    #[test]
    fn test_get_row_empty() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "get_empty.xlsx");

        let result = get(&file_path, "Sheet1", 1, &default_global());
        assert!(result.is_ok());
    }

    #[test]
    fn test_append_row_json() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "append_json.xlsx");

        let mut global = default_global();
        global.format = OutputFormat::Json;
        global.quiet = false;

        let result = append(&file_path, "Sheet1", "1,2,3", &global);
        assert!(result.is_ok());
    }

    #[test]
    fn test_insert_row_json() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "insert_json.xlsx");
        setup_test_data(&file_path);

        let mut global = default_global();
        global.format = OutputFormat::Json;
        global.quiet = false;

        let result = insert(&file_path, "Sheet1", 2, &global);
        assert!(result.is_ok());
    }

    #[test]
    fn test_delete_row_json() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "delete_json.xlsx");
        setup_test_data(&file_path);

        let mut global = default_global();
        global.format = OutputFormat::Json;
        global.quiet = false;

        let result = delete(&file_path, "Sheet1", 1, &global);
        assert!(result.is_ok());
    }

    #[test]
    fn test_copy_row_json() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "copy_json.xlsx");
        setup_test_data(&file_path);

        let mut global = default_global();
        global.format = OutputFormat::Json;
        global.quiet = false;

        let result = copy(&file_path, "Sheet1", 1, 5, &global);
        assert!(result.is_ok());
    }

    #[test]
    fn test_height_json() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "height_json.xlsx");

        let mut global = default_global();
        global.format = OutputFormat::Json;
        global.quiet = false;

        let result = height(&file_path, "Sheet1", 1, Some(30.0), &global);
        assert!(result.is_ok());
    }

    #[test]
    fn test_height_get_json() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "height_get_json.xlsx");

        height(&file_path, "Sheet1", 1, Some(30.0), &default_global()).unwrap();

        let mut global = default_global();
        global.format = OutputFormat::Json;

        let result = height(&file_path, "Sheet1", 1, None, &global);
        assert!(result.is_ok());
    }

    #[test]
    fn test_find_rows_json() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "find_json.xlsx");
        setup_test_data(&file_path);

        let mut global = default_global();
        global.format = OutputFormat::Json;

        let result = find(&file_path, "Sheet1", "1", None, &global);
        assert!(result.is_ok());
    }

    #[test]
    fn test_find_rows_no_match() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "find_nomatch.xlsx");
        setup_test_data(&file_path);

        let result = find(&file_path, "Sheet1", "nonexistent", None, &default_global());
        assert!(result.is_ok());
    }

    #[test]
    fn test_sheet_not_found() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "notfound.xlsx");

        let result = get(&file_path, "NonExistent", 1, &default_global());
        assert!(result.is_err());
    }

    #[test]
    fn test_append_with_string_values() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "append_str.xlsx");

        let result = append(&file_path, "Sheet1", "hello,world,test", &default_global());
        assert!(result.is_ok());

        let wb = Workbook::open(&file_path).unwrap();
        let cell_ref = xlex_core::CellRef::new(1, 1);
        let value = wb.get_cell("Sheet1", &cell_ref).unwrap();
        assert_eq!(value, CellValue::String("hello".to_string()));
    }

    #[test]
    fn test_insert_verbose() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "insert_verbose.xlsx");
        setup_test_data(&file_path);

        let mut global = default_global();
        global.quiet = false;

        let result = insert(&file_path, "Sheet1", 2, &global);
        assert!(result.is_ok());
    }

    #[test]
    fn test_delete_verbose() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "delete_verbose.xlsx");
        setup_test_data(&file_path);

        let mut global = default_global();
        global.quiet = false;

        let result = delete(&file_path, "Sheet1", 1, &global);
        assert!(result.is_ok());
    }

    #[test]
    fn test_copy_verbose() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "copy_verbose.xlsx");
        setup_test_data(&file_path);

        let mut global = default_global();
        global.quiet = false;

        let result = copy(&file_path, "Sheet1", 1, 5, &global);
        assert!(result.is_ok());
    }

    #[test]
    fn test_move_verbose() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "move_verbose.xlsx");
        setup_test_data(&file_path);

        let mut global = default_global();
        global.quiet = false;

        let result = move_row(&file_path, "Sheet1", 1, 5, &global);
        assert!(result.is_ok());
    }

    #[test]
    fn test_hide_verbose() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "hide_verbose.xlsx");

        let mut global = default_global();
        global.quiet = false;

        let result = hide(&file_path, "Sheet1", 1, &global);
        assert!(result.is_ok());
    }

    #[test]
    fn test_unhide_verbose() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "unhide_verbose.xlsx");

        // First hide the row
        hide(&file_path, "Sheet1", 1, &default_global()).unwrap();

        let mut global = default_global();
        global.quiet = false;

        let result = unhide(&file_path, "Sheet1", 1, &global);
        assert!(result.is_ok());
    }
}
