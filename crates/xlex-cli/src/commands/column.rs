//! Column operations.

use anyhow::Result;
use clap::{Parser, Subcommand};
use colored::Colorize;

use xlex_core::{CellRef, Workbook};

use super::{GlobalOptions, OutputFormat};

/// Arguments for column operations.
#[derive(Parser)]
pub struct ColumnArgs {
    #[command(subcommand)]
    pub command: ColumnCommand,
}

#[derive(Subcommand)]
pub enum ColumnCommand {
    /// Get column data
    Get {
        /// Path to the xlsx file
        file: std::path::PathBuf,
        /// Sheet name
        sheet: String,
        /// Column (letter, e.g., A, B, AA)
        column: String,
    },
    /// Insert a column
    Insert {
        /// Path to the xlsx file
        file: std::path::PathBuf,
        /// Sheet name
        sheet: String,
        /// Column to insert at
        column: String,
    },
    /// Delete a column
    Delete {
        /// Path to the xlsx file
        file: std::path::PathBuf,
        /// Sheet name
        sheet: String,
        /// Column to delete
        column: String,
    },
    /// Copy a column
    Copy {
        /// Path to the xlsx file
        file: std::path::PathBuf,
        /// Sheet name
        sheet: String,
        /// Source column
        source: String,
        /// Destination column
        dest: String,
    },
    /// Move a column
    Move {
        /// Path to the xlsx file
        file: std::path::PathBuf,
        /// Sheet name
        sheet: String,
        /// Source column
        source: String,
        /// Destination column
        dest: String,
    },
    /// Set column width
    Width {
        /// Path to the xlsx file
        file: std::path::PathBuf,
        /// Sheet name
        sheet: String,
        /// Column
        column: String,
        /// Width in characters (omit to show current)
        width: Option<f64>,
    },
    /// Hide a column
    Hide {
        /// Path to the xlsx file
        file: std::path::PathBuf,
        /// Sheet name
        sheet: String,
        /// Column
        column: String,
    },
    /// Unhide a column
    Unhide {
        /// Path to the xlsx file
        file: std::path::PathBuf,
        /// Sheet name
        sheet: String,
        /// Column
        column: String,
    },
    /// Get column header (first row value)
    Header {
        /// Path to the xlsx file
        file: std::path::PathBuf,
        /// Sheet name
        sheet: String,
        /// Column
        column: String,
    },
    /// Find columns matching criteria
    Find {
        /// Path to the xlsx file
        file: std::path::PathBuf,
        /// Sheet name
        sheet: String,
        /// Search pattern
        pattern: String,
    },
    /// Column statistics
    Stats {
        /// Path to the xlsx file
        file: std::path::PathBuf,
        /// Sheet name
        sheet: String,
        /// Column
        column: String,
    },
}

/// Run column operations.
pub fn run(args: &ColumnArgs, global: &GlobalOptions) -> Result<()> {
    match &args.command {
        ColumnCommand::Get {
            file,
            sheet,
            column,
        } => get(file, sheet, column, global),
        ColumnCommand::Insert {
            file,
            sheet,
            column,
        } => insert(file, sheet, column, global),
        ColumnCommand::Delete {
            file,
            sheet,
            column,
        } => delete(file, sheet, column, global),
        ColumnCommand::Copy {
            file,
            sheet,
            source,
            dest,
        } => copy(file, sheet, source, dest, global),
        ColumnCommand::Move {
            file,
            sheet,
            source,
            dest,
        } => move_column(file, sheet, source, dest, global),
        ColumnCommand::Width {
            file,
            sheet,
            column,
            width: col_width,
        } => width(file, sheet, column, *col_width, global),
        ColumnCommand::Hide {
            file,
            sheet,
            column,
        } => hide(file, sheet, column, global),
        ColumnCommand::Unhide {
            file,
            sheet,
            column,
        } => unhide(file, sheet, column, global),
        ColumnCommand::Header {
            file,
            sheet,
            column,
        } => header(file, sheet, column, global),
        ColumnCommand::Find {
            file,
            sheet,
            pattern,
        } => find(file, sheet, pattern, global),
        ColumnCommand::Stats {
            file,
            sheet,
            column,
        } => stats(file, sheet, column, global),
    }
}

fn parse_column(col: &str) -> Result<u32> {
    CellRef::col_from_letters_pub(&col.to_uppercase())
        .ok_or_else(|| anyhow::anyhow!("Invalid column: {}", col))
}

fn get(file: &std::path::Path, sheet: &str, column: &str, global: &GlobalOptions) -> Result<()> {
    let workbook = Workbook::open(file)?;
    let col = parse_column(column)?;
    let sheet_obj =
        workbook
            .get_sheet(sheet)
            .ok_or_else(|| xlex_core::XlexError::SheetNotFound {
                name: sheet.to_string(),
            })?;

    // Collect cells in this column
    let mut col_cells: Vec<_> = sheet_obj
        .cells()
        .filter(|c| c.reference.col == col)
        .collect();
    col_cells.sort_by_key(|c| c.reference.row);

    if global.format == OutputFormat::Json {
        let values: Vec<_> = col_cells
            .iter()
            .map(|c| {
                serde_json::json!({
                    "row": c.reference.row,
                    "ref": c.reference.to_a1(),
                    "value": c.value.to_display_string(),
                    "type": c.value.type_name(),
                })
            })
            .collect();
        println!("{}", serde_json::to_string_pretty(&values)?);
    } else {
        for cell in col_cells {
            println!("{}: {}", cell.reference.to_a1().cyan(), cell.value);
        }
    }

    Ok(())
}

fn insert(file: &std::path::Path, sheet: &str, column: &str, global: &GlobalOptions) -> Result<()> {
    let col = parse_column(column)?;

    if global.dry_run {
        println!("Would insert column at {}", column);
        return Ok(());
    }

    let mut workbook = Workbook::open(file)?;
    let sheet_obj =
        workbook
            .get_sheet_mut(sheet)
            .ok_or_else(|| xlex_core::XlexError::SheetNotFound {
                name: sheet.to_string(),
            })?;

    sheet_obj.insert_columns(col, 1);
    let _ = sheet_obj;
    workbook.save()?;

    if !global.quiet {
        println!("Inserted column at {}", column.to_uppercase().green());
    }

    Ok(())
}

fn delete(file: &std::path::Path, sheet: &str, column: &str, global: &GlobalOptions) -> Result<()> {
    let col = parse_column(column)?;

    if global.dry_run {
        println!("Would delete column {}", column);
        return Ok(());
    }

    let mut workbook = Workbook::open(file)?;
    let sheet_obj =
        workbook
            .get_sheet_mut(sheet)
            .ok_or_else(|| xlex_core::XlexError::SheetNotFound {
                name: sheet.to_string(),
            })?;

    sheet_obj.delete_columns(col, 1);
    let _ = sheet_obj;
    workbook.save()?;

    if !global.quiet {
        println!("Deleted column {}", column.to_uppercase().green());
    }

    Ok(())
}

fn copy(
    file: &std::path::Path,
    sheet: &str,
    source: &str,
    dest: &str,
    global: &GlobalOptions,
) -> Result<()> {
    let source_col = parse_column(source)?;
    let dest_col = parse_column(dest)?;

    if global.dry_run {
        println!("Would copy column {} to column {}", source, dest);
        return Ok(());
    }

    let mut workbook = Workbook::open(file)?;

    // Get source column data
    let sheet_obj =
        workbook
            .get_sheet(sheet)
            .ok_or_else(|| xlex_core::XlexError::SheetNotFound {
                name: sheet.to_string(),
            })?;

    // Collect all cells in the source column
    let source_cells: Vec<_> = sheet_obj
        .cells()
        .filter(|c| c.reference.col == source_col)
        .map(|c| (c.reference.row, c.value.clone()))
        .collect();

    let source_width = sheet_obj.get_column_width(source_col);
    let _ = sheet_obj;

    // Insert a new column at destination
    let sheet_obj =
        workbook
            .get_sheet_mut(sheet)
            .ok_or_else(|| xlex_core::XlexError::SheetNotFound {
                name: sheet.to_string(),
            })?;

    sheet_obj.insert_columns(dest_col, 1);

    // Copy cells to destination
    for (row, value) in source_cells {
        let dest_ref = xlex_core::CellRef::new(dest_col, row);
        sheet_obj.set_cell(dest_ref, value);
    }

    // Copy column width if set
    if let Some(w) = source_width {
        sheet_obj.set_column_width(dest_col, w);
    }

    let _ = sheet_obj;
    workbook.save()?;

    if !global.quiet {
        println!(
            "Copied column {} to column {}",
            source.to_uppercase().cyan(),
            dest.to_uppercase().green()
        );
    }

    Ok(())
}

fn move_column(
    file: &std::path::Path,
    sheet: &str,
    source: &str,
    dest: &str,
    global: &GlobalOptions,
) -> Result<()> {
    let source_col = parse_column(source)?;
    let dest_col = parse_column(dest)?;

    if global.dry_run {
        println!("Would move column {} to column {}", source, dest);
        return Ok(());
    }

    let mut workbook = Workbook::open(file)?;

    // Get source column data
    let sheet_obj =
        workbook
            .get_sheet(sheet)
            .ok_or_else(|| xlex_core::XlexError::SheetNotFound {
                name: sheet.to_string(),
            })?;

    // Collect all cells in the source column
    let source_cells: Vec<_> = sheet_obj
        .cells()
        .filter(|c| c.reference.col == source_col)
        .map(|c| (c.reference.row, c.value.clone()))
        .collect();

    let source_width = sheet_obj.get_column_width(source_col);
    let _ = sheet_obj;

    let sheet_obj =
        workbook
            .get_sheet_mut(sheet)
            .ok_or_else(|| xlex_core::XlexError::SheetNotFound {
                name: sheet.to_string(),
            })?;

    // The logic differs based on whether moving left or right
    if source_col < dest_col {
        // Moving right: insert first, then delete (adjusted source)
        sheet_obj.insert_columns(dest_col + 1, 1);
        let actual_dest = dest_col + 1;

        // Copy cells
        for (row, value) in &source_cells {
            let dest_ref = xlex_core::CellRef::new(actual_dest, *row);
            sheet_obj.set_cell(dest_ref, value.clone());
        }

        if let Some(w) = source_width {
            sheet_obj.set_column_width(actual_dest, w);
        }

        // Delete original column
        sheet_obj.delete_columns(source_col, 1);
    } else {
        // Moving left: delete first, then insert at adjusted position
        sheet_obj.delete_columns(source_col, 1);
        sheet_obj.insert_columns(dest_col, 1);

        // Copy cells
        for (row, value) in &source_cells {
            let dest_ref = xlex_core::CellRef::new(dest_col, *row);
            sheet_obj.set_cell(dest_ref, value.clone());
        }

        if let Some(w) = source_width {
            sheet_obj.set_column_width(dest_col, w);
        }
    }

    let _ = sheet_obj;
    workbook.save()?;

    if !global.quiet {
        println!(
            "Moved column {} to column {}",
            source.to_uppercase().cyan(),
            dest.to_uppercase().green()
        );
    }

    Ok(())
}

fn width(
    file: &std::path::Path,
    sheet: &str,
    column: &str,
    width: Option<f64>,
    global: &GlobalOptions,
) -> Result<()> {
    let col = parse_column(column)?;

    if let Some(w) = width {
        if global.dry_run {
            println!("Would set column {} width to {}", column, w);
            return Ok(());
        }

        let mut workbook = Workbook::open(file)?;
        let sheet_obj =
            workbook
                .get_sheet_mut(sheet)
                .ok_or_else(|| xlex_core::XlexError::SheetNotFound {
                    name: sheet.to_string(),
                })?;

        sheet_obj.set_column_width(col, w);
        let _ = sheet_obj;
        workbook.save()?;

        if !global.quiet {
            println!("Set column {} width to {}", column.to_uppercase().cyan(), w);
        }
    } else {
        let workbook = Workbook::open(file)?;
        let sheet_obj =
            workbook
                .get_sheet(sheet)
                .ok_or_else(|| xlex_core::XlexError::SheetNotFound {
                    name: sheet.to_string(),
                })?;

        if let Some(w) = sheet_obj.get_column_width(col) {
            println!("{}", w);
        } else {
            println!("default");
        }
    }

    Ok(())
}

fn hide(file: &std::path::Path, sheet: &str, column: &str, global: &GlobalOptions) -> Result<()> {
    if global.dry_run {
        println!("Would hide column {}", column);
        return Ok(());
    }

    let col = parse_column(column)?;
    let mut workbook = Workbook::open(file)?;
    let sheet_obj =
        workbook
            .get_sheet_mut(sheet)
            .ok_or_else(|| xlex_core::XlexError::SheetNotFound {
                name: sheet.to_string(),
            })?;

    sheet_obj.set_column_hidden(col, true);
    let _ = sheet_obj;
    workbook.save()?;

    if !global.quiet {
        println!("Hid column {}", column.to_uppercase().dimmed());
    }

    Ok(())
}

fn unhide(file: &std::path::Path, sheet: &str, column: &str, global: &GlobalOptions) -> Result<()> {
    if global.dry_run {
        println!("Would unhide column {}", column);
        return Ok(());
    }

    let col = parse_column(column)?;
    let mut workbook = Workbook::open(file)?;
    let sheet_obj =
        workbook
            .get_sheet_mut(sheet)
            .ok_or_else(|| xlex_core::XlexError::SheetNotFound {
                name: sheet.to_string(),
            })?;

    sheet_obj.set_column_hidden(col, false);
    let _ = sheet_obj;
    workbook.save()?;

    if !global.quiet {
        println!("Unhid column {}", column.to_uppercase().green());
    }

    Ok(())
}

fn header(file: &std::path::Path, sheet: &str, column: &str, global: &GlobalOptions) -> Result<()> {
    let workbook = Workbook::open(file)?;
    let col = parse_column(column)?;
    let cell_ref = CellRef::new(col, 1);
    let value = workbook.get_cell(sheet, &cell_ref)?;

    if global.format == OutputFormat::Json {
        let json = serde_json::json!({
            "column": column.to_uppercase(),
            "header": value.to_display_string(),
        });
        println!("{}", serde_json::to_string_pretty(&json)?);
    } else {
        println!("{}", value.to_display_string());
    }

    Ok(())
}

fn find(file: &std::path::Path, sheet: &str, pattern: &str, global: &GlobalOptions) -> Result<()> {
    let workbook = Workbook::open(file)?;
    let sheet_obj =
        workbook
            .get_sheet(sheet)
            .ok_or_else(|| xlex_core::XlexError::SheetNotFound {
                name: sheet.to_string(),
            })?;

    let mut matches: Vec<u32> = Vec::new();
    for cell in sheet_obj.cells() {
        let value = cell.value.to_display_string();
        if value.contains(pattern) && !matches.contains(&cell.reference.col) {
            matches.push(cell.reference.col);
        }
    }

    matches.sort();
    let col_letters: Vec<String> = matches
        .iter()
        .map(|c| CellRef::col_to_letters(*c))
        .collect();

    if global.format == OutputFormat::Json {
        let json = serde_json::json!({
            "pattern": pattern,
            "matches": col_letters,
            "count": matches.len(),
        });
        println!("{}", serde_json::to_string_pretty(&json)?);
    } else {
        for col in col_letters {
            println!("{}", col);
        }
    }

    Ok(())
}

fn stats(file: &std::path::Path, sheet: &str, column: &str, global: &GlobalOptions) -> Result<()> {
    let workbook = Workbook::open(file)?;
    let col = parse_column(column)?;
    let sheet_obj =
        workbook
            .get_sheet(sheet)
            .ok_or_else(|| xlex_core::XlexError::SheetNotFound {
                name: sheet.to_string(),
            })?;

    // Collect statistics
    let mut count = 0;
    let mut numeric_count = 0;
    let mut sum = 0.0;
    let mut min: Option<f64> = None;
    let mut max: Option<f64> = None;

    for cell in sheet_obj.cells() {
        if cell.reference.col != col {
            continue;
        }
        count += 1;

        if let xlex_core::CellValue::Number(n) = cell.value {
            numeric_count += 1;
            sum += n;
            min = Some(min.map_or(n, |m| m.min(n)));
            max = Some(max.map_or(n, |m| m.max(n)));
        }
    }

    let avg = if numeric_count > 0 {
        Some(sum / numeric_count as f64)
    } else {
        None
    };

    if global.format == OutputFormat::Json {
        let json = serde_json::json!({
            "column": column.to_uppercase(),
            "count": count,
            "numericCount": numeric_count,
            "sum": if numeric_count > 0 { Some(sum) } else { None },
            "average": avg,
            "min": min,
            "max": max,
        });
        println!("{}", serde_json::to_string_pretty(&json)?);
    } else {
        println!("{}: {}", "Column".bold(), column.to_uppercase());
        println!("  {}: {}", "Count".cyan(), count);
        println!("  {}: {}", "Numeric".cyan(), numeric_count);
        if numeric_count > 0 {
            println!("  {}: {}", "Sum".cyan(), sum);
            if let Some(a) = avg {
                println!("  {}: {:.2}", "Average".cyan(), a);
            }
            if let Some(m) = min {
                println!("  {}: {}", "Min".cyan(), m);
            }
            if let Some(m) = max {
                println!("  {}: {}", "Max".cyan(), m);
            }
        }
    }

    Ok(())
}
