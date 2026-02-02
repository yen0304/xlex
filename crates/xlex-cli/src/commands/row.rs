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
        RowCommand::Append { file, sheet, values } => append(file, sheet, values, global),
        RowCommand::Insert { file, sheet, row } => insert(file, sheet, *row, global),
        RowCommand::Delete { file, sheet, row } => delete(file, sheet, *row, global),
        RowCommand::Copy { file, sheet, source, dest } => copy(file, sheet, *source, *dest, global),
        RowCommand::Move { file, sheet, source, dest } => move_row(file, sheet, *source, *dest, global),
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
    let sheet_obj = workbook
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
        let values: Vec<_> = row_cells.iter().map(|c| c.value.to_display_string()).collect();
        println!("{}", values.join(","));
    } else {
        for cell in row_cells {
            println!("{}: {}", cell.reference.to_a1().cyan(), cell.value);
        }
    }

    Ok(())
}

fn append(
    file: &std::path::Path,
    sheet: &str,
    values: &str,
    global: &GlobalOptions,
) -> Result<()> {
    if global.dry_run {
        println!("Would append row with values: {}", values);
        return Ok(());
    }

    let mut workbook = Workbook::open(file)?;
    let sheet_obj = workbook
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
    let sheet_obj = workbook
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
    let sheet_obj = workbook
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
    let sheet_obj = workbook
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
    let sheet_obj = workbook
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
        println!("Copied row {} to row {}", source.to_string().cyan(), dest.to_string().green());
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
    let sheet_obj = workbook
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

    let sheet_obj = workbook
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
        println!("Moved row {} to row {}", source.to_string().cyan(), dest.to_string().green());
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
        let sheet_obj = workbook
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
        let sheet_obj = workbook
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
    let sheet_obj = workbook
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
    let sheet_obj = workbook
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
    let sheet_obj = workbook
        .get_sheet(sheet)
        .ok_or_else(|| xlex_core::XlexError::SheetNotFound {
            name: sheet.to_string(),
        })?;

    let col_filter: Option<u32> = column.and_then(|c| {
        xlex_core::CellRef::col_from_letters_pub(&c.to_uppercase())
    });

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

