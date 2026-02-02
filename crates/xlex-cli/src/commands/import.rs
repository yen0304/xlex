//! Import operations.

use anyhow::Result;
use clap::{Parser, Subcommand, ValueEnum};
use colored::Colorize;

use xlex_core::{CellRef, CellValue, Workbook};

use super::GlobalOptions;

/// Arguments for import operations.
#[derive(Parser)]
pub struct ImportArgs {
    #[command(subcommand)]
    pub command: ImportCommand,
}

#[derive(Clone, ValueEnum)]
pub enum ImportFormat {
    Csv,
    Json,
    Tsv,
}

#[derive(Subcommand)]
pub enum ImportCommand {
    /// Import CSV file
    Csv {
        /// Source CSV file
        source: std::path::PathBuf,
        /// Destination xlsx file
        dest: std::path::PathBuf,
        /// Sheet name (default: Sheet1)
        #[arg(short, long)]
        sheet: Option<String>,
        /// Delimiter character
        #[arg(short, long, default_value = ",")]
        delimiter: char,
        /// Has header row
        #[arg(long)]
        header: bool,
    },
    /// Import JSON file
    Json {
        /// Source JSON file
        source: std::path::PathBuf,
        /// Destination xlsx file
        dest: std::path::PathBuf,
        /// Sheet name (default: Sheet1)
        #[arg(short, long)]
        sheet: Option<String>,
    },
    /// Import TSV file
    Tsv {
        /// Source TSV file
        source: std::path::PathBuf,
        /// Destination xlsx file
        dest: std::path::PathBuf,
        /// Sheet name (default: Sheet1)
        #[arg(short, long)]
        sheet: Option<String>,
    },
    /// Import NDJSON file (newline-delimited JSON)
    Ndjson {
        /// Source NDJSON file
        source: std::path::PathBuf,
        /// Destination xlsx file
        dest: std::path::PathBuf,
        /// Sheet name (default: Sheet1)
        #[arg(short, long)]
        sheet: Option<String>,
        /// Use first row keys as headers
        #[arg(long)]
        header: bool,
    },
}

/// Run import operations.
pub fn run(args: &ImportArgs, global: &GlobalOptions) -> Result<()> {
    match &args.command {
        ImportCommand::Csv {
            source,
            dest,
            sheet,
            delimiter,
            header,
        } => import_csv(source, dest, sheet.as_deref(), *delimiter, *header, global),
        ImportCommand::Json {
            source,
            dest,
            sheet,
        } => import_json(source, dest, sheet.as_deref(), global),
        ImportCommand::Tsv {
            source,
            dest,
            sheet,
        } => import_tsv(source, dest, sheet.as_deref(), global),
        ImportCommand::Ndjson {
            source,
            dest,
            sheet,
            header,
        } => import_ndjson(source, dest, sheet.as_deref(), *header, global),
    }
}

fn import_csv(
    source: &std::path::Path,
    dest: &std::path::Path,
    sheet: Option<&str>,
    delimiter: char,
    _has_header: bool,
    global: &GlobalOptions,
) -> Result<()> {
    if global.dry_run {
        println!("Would import {} to {}", source.display(), dest.display());
        return Ok(());
    }

    let sheet_name = sheet.unwrap_or("Sheet1");
    let content = std::fs::read_to_string(source)?;

    let mut workbook = if dest.exists() {
        Workbook::open(dest)?
    } else {
        Workbook::with_sheets(&[sheet_name])
    };

    // Ensure sheet exists
    if workbook.get_sheet(sheet_name).is_none() {
        workbook.add_sheet(sheet_name)?;
    }

    let mut row = 1u32;
    for line in content.lines() {
        let mut col = 1u32;
        for value in line.split(delimiter) {
            let cell_ref = CellRef::new(col, row);
            let cell_value = parse_value(value.trim());
            workbook.set_cell(sheet_name, cell_ref, cell_value)?;
            col += 1;
        }
        row += 1;
    }

    workbook.save_as(dest)?;

    if !global.quiet {
        println!(
            "Imported {} rows to {}",
            (row - 1).to_string().green(),
            dest.display()
        );
    }

    Ok(())
}

fn import_json(
    source: &std::path::Path,
    dest: &std::path::Path,
    sheet: Option<&str>,
    global: &GlobalOptions,
) -> Result<()> {
    if global.dry_run {
        println!("Would import {} to {}", source.display(), dest.display());
        return Ok(());
    }

    let sheet_name = sheet.unwrap_or("Sheet1");
    let content = std::fs::read_to_string(source)?;
    let data: serde_json::Value = serde_json::from_str(&content)?;

    let mut workbook = if dest.exists() {
        Workbook::open(dest)?
    } else {
        Workbook::with_sheets(&[sheet_name])
    };

    // Ensure sheet exists
    if workbook.get_sheet(sheet_name).is_none() {
        workbook.add_sheet(sheet_name)?;
    }

    match data {
        serde_json::Value::Array(arr) => {
            // Array of objects or arrays
            if let Some(first) = arr.first() {
                if first.is_object() {
                    // Array of objects - use keys as headers
                    if let serde_json::Value::Object(obj) = first {
                        let keys: Vec<_> = obj.keys().collect();

                        // Write headers
                        for (col, key) in keys.iter().enumerate() {
                            let cell_ref = CellRef::new((col + 1) as u32, 1);
                            workbook.set_cell(
                                sheet_name,
                                cell_ref,
                                CellValue::String((*key).clone()),
                            )?;
                        }

                        // Write data
                        for (row, item) in arr.iter().enumerate() {
                            if let serde_json::Value::Object(obj) = item {
                                for (col, key) in keys.iter().enumerate() {
                                    let cell_ref = CellRef::new((col + 1) as u32, (row + 2) as u32);
                                    let value = obj.get(*key).unwrap_or(&serde_json::Value::Null);
                                    workbook.set_cell(sheet_name, cell_ref, json_to_cell(value))?;
                                }
                            }
                        }
                    }
                } else if first.is_array() {
                    // Array of arrays
                    for (row, item) in arr.iter().enumerate() {
                        if let serde_json::Value::Array(row_arr) = item {
                            for (col, value) in row_arr.iter().enumerate() {
                                let cell_ref = CellRef::new((col + 1) as u32, (row + 1) as u32);
                                workbook.set_cell(sheet_name, cell_ref, json_to_cell(value))?;
                            }
                        }
                    }
                }
            }
        }
        _ => {
            anyhow::bail!("JSON must be an array of objects or arrays");
        }
    }

    workbook.save_as(dest)?;

    if !global.quiet {
        println!("Imported JSON data to {}", dest.display());
    }

    Ok(())
}

fn import_tsv(
    source: &std::path::Path,
    dest: &std::path::Path,
    sheet: Option<&str>,
    global: &GlobalOptions,
) -> Result<()> {
    import_csv(source, dest, sheet, '\t', false, global)
}

fn parse_value(s: &str) -> CellValue {
    if s.is_empty() {
        return CellValue::Empty;
    }

    // Try number
    if let Ok(n) = s.parse::<f64>() {
        return CellValue::Number(n);
    }

    // Try boolean
    match s.to_lowercase().as_str() {
        "true" | "yes" | "1" => return CellValue::Boolean(true),
        "false" | "no" | "0" if s != "0" => return CellValue::Boolean(false),
        _ => {}
    }

    CellValue::String(s.to_string())
}

fn json_to_cell(value: &serde_json::Value) -> CellValue {
    match value {
        serde_json::Value::Null => CellValue::Empty,
        serde_json::Value::Bool(b) => CellValue::Boolean(*b),
        serde_json::Value::Number(n) => {
            CellValue::Number(n.as_f64().unwrap_or(0.0))
        }
        serde_json::Value::String(s) => CellValue::String(s.clone()),
        _ => CellValue::String(value.to_string()),
    }
}

fn import_ndjson(
    source: &std::path::Path,
    dest: &std::path::Path,
    sheet: Option<&str>,
    has_header: bool,
    global: &GlobalOptions,
) -> Result<()> {
    if global.dry_run {
        println!("Would import {} to {}", source.display(), dest.display());
        return Ok(());
    }

    let sheet_name = sheet.unwrap_or("Sheet1");
    let content = std::fs::read_to_string(source)?;
    let lines: Vec<&str> = content.lines().filter(|l| !l.trim().is_empty()).collect();

    if lines.is_empty() {
        anyhow::bail!("NDJSON file is empty");
    }

    let mut workbook = if dest.exists() {
        Workbook::open(dest)?
    } else {
        Workbook::with_sheets(&[sheet_name])
    };

    // Ensure sheet exists
    if workbook.get_sheet(sheet_name).is_none() {
        workbook.add_sheet(sheet_name)?;
    }

    // Parse first line to determine format
    let first: serde_json::Value = serde_json::from_str(lines[0])?;

    if first.is_object() && has_header {
        // Object format - use keys as headers
        if let serde_json::Value::Object(obj) = &first {
            let keys: Vec<_> = obj.keys().cloned().collect();

            // Write headers
            for (col, key) in keys.iter().enumerate() {
                let cell_ref = CellRef::new((col + 1) as u32, 1);
                workbook.set_cell(sheet_name, cell_ref, CellValue::String(key.clone()))?;
            }

            // Write data
            for (row_idx, line) in lines.iter().enumerate() {
                let item: serde_json::Value = serde_json::from_str(line)?;
                if let serde_json::Value::Object(obj) = item {
                    for (col, key) in keys.iter().enumerate() {
                        let cell_ref = CellRef::new((col + 1) as u32, (row_idx + 2) as u32);
                        let value = obj.get(key).unwrap_or(&serde_json::Value::Null);
                        workbook.set_cell(sheet_name, cell_ref, json_to_cell(value))?;
                    }
                }
            }
        }
    } else if first.is_array() {
        // Array format
        for (row_idx, line) in lines.iter().enumerate() {
            let item: serde_json::Value = serde_json::from_str(line)?;
            if let serde_json::Value::Array(arr) = item {
                for (col, value) in arr.iter().enumerate() {
                    let cell_ref = CellRef::new((col + 1) as u32, (row_idx + 1) as u32);
                    workbook.set_cell(sheet_name, cell_ref, json_to_cell(value))?;
                }
            }
        }
    } else if first.is_object() {
        // Object format without headers
        for (row_idx, line) in lines.iter().enumerate() {
            let item: serde_json::Value = serde_json::from_str(line)?;
            if let serde_json::Value::Object(obj) = item {
                for (col, (_key, value)) in obj.iter().enumerate() {
                    let cell_ref = CellRef::new((col + 1) as u32, (row_idx + 1) as u32);
                    workbook.set_cell(sheet_name, cell_ref, json_to_cell(value))?;
                }
            }
        }
    }

    workbook.save_as(dest)?;

    if !global.quiet {
        println!(
            "Imported {} rows to {}",
            lines.len().to_string().green(),
            dest.display()
        );
    }

    Ok(())
}
