//! Import operations.

use anyhow::Result;
use clap::{Parser, Subcommand, ValueEnum};
use colored::Colorize;

use xlex_core::{CellRef, CellValue, Workbook};

use super::GlobalOptions;
use crate::progress::Progress;

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
    let lines: Vec<&str> = content.lines().collect();
    let total_lines = lines.len();

    // Create progress bar for large files
    let progress = if total_lines > 100 {
        Some(Progress::bar(
            total_lines as u64,
            "Importing CSV...",
            global.quiet,
        ))
    } else {
        None
    };

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
    for line in lines.iter() {
        let mut col = 1u32;
        for value in line.split(delimiter) {
            let cell_ref = CellRef::new(col, row);
            let cell_value = parse_value(value.trim());
            workbook.set_cell(sheet_name, cell_ref, cell_value)?;
            col += 1;
        }
        row += 1;
        if let Some(ref pb) = progress {
            pb.inc(1);
        }
    }

    if let Some(ref pb) = progress {
        pb.finish_and_clear();
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
        serde_json::Value::Number(n) => CellValue::Number(n.as_f64().unwrap_or(0.0)),
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

#[cfg(test)]
mod tests {
    use super::*;
    use tempfile::TempDir;

    fn default_global() -> GlobalOptions {
        GlobalOptions {
            quiet: true,
            verbose: false,
            format: super::super::OutputFormat::Text,
            no_color: true,
            color: false,
            json_errors: false,
            dry_run: false,
            output: None,
        }
    }

    #[test]
    fn test_parse_value_number() {
        assert_eq!(parse_value("42"), CellValue::Number(42.0));
        assert_eq!(parse_value("3.14"), CellValue::Number(3.14));
        assert_eq!(parse_value("-100"), CellValue::Number(-100.0));
    }

    #[test]
    fn test_parse_value_boolean() {
        assert_eq!(parse_value("true"), CellValue::Boolean(true));
        assert_eq!(parse_value("TRUE"), CellValue::Boolean(true));
        assert_eq!(parse_value("yes"), CellValue::Boolean(true));
        assert_eq!(parse_value("false"), CellValue::Boolean(false));
    }

    #[test]
    fn test_parse_value_string() {
        assert_eq!(parse_value("hello"), CellValue::String("hello".to_string()));
    }

    #[test]
    fn test_parse_value_empty() {
        assert_eq!(parse_value(""), CellValue::Empty);
    }

    #[test]
    fn test_json_to_cell() {
        assert_eq!(json_to_cell(&serde_json::Value::Null), CellValue::Empty);
        assert_eq!(
            json_to_cell(&serde_json::Value::Bool(true)),
            CellValue::Boolean(true)
        );
        assert_eq!(
            json_to_cell(&serde_json::json!(42)),
            CellValue::Number(42.0)
        );
        assert_eq!(
            json_to_cell(&serde_json::Value::String("test".to_string())),
            CellValue::String("test".to_string())
        );
    }

    #[test]
    fn test_import_csv() {
        let temp_dir = TempDir::new().unwrap();
        let csv_path = temp_dir.path().join("data.csv");
        let xlsx_path = temp_dir.path().join("output.xlsx");

        std::fs::write(&csv_path, "Name,Age\nAlice,30\nBob,25").unwrap();

        let result = import_csv(&csv_path, &xlsx_path, None, ',', false, &default_global());
        assert!(result.is_ok());
        assert!(xlsx_path.exists());

        let wb = Workbook::open(&xlsx_path).unwrap();
        let cell_ref = CellRef::new(1, 1);
        let value = wb.get_cell("Sheet1", &cell_ref).unwrap();
        assert_eq!(value, CellValue::String("Name".to_string()));
    }

    #[test]
    fn test_import_csv_dry_run() {
        let temp_dir = TempDir::new().unwrap();
        let csv_path = temp_dir.path().join("data.csv");
        let xlsx_path = temp_dir.path().join("output.xlsx");

        std::fs::write(&csv_path, "Name,Age\nAlice,30").unwrap();

        let mut global = default_global();
        global.dry_run = true;

        let result = import_csv(&csv_path, &xlsx_path, None, ',', false, &global);
        assert!(result.is_ok());
        assert!(!xlsx_path.exists()); // Should not create file
    }

    #[test]
    fn test_import_tsv() {
        let temp_dir = TempDir::new().unwrap();
        let tsv_path = temp_dir.path().join("data.tsv");
        let xlsx_path = temp_dir.path().join("output.xlsx");

        std::fs::write(&tsv_path, "Name\tAge\nAlice\t30\nBob\t25").unwrap();

        let result = import_tsv(&tsv_path, &xlsx_path, None, &default_global());
        assert!(result.is_ok());
        assert!(xlsx_path.exists());
    }

    #[test]
    fn test_import_json_array_of_objects() {
        let temp_dir = TempDir::new().unwrap();
        let json_path = temp_dir.path().join("data.json");
        let xlsx_path = temp_dir.path().join("output.xlsx");

        std::fs::write(
            &json_path,
            r#"[{"name": "Alice", "age": 30}, {"name": "Bob", "age": 25}]"#,
        )
        .unwrap();

        let result = import_json(&json_path, &xlsx_path, None, &default_global());
        assert!(result.is_ok());
        assert!(xlsx_path.exists());
    }

    #[test]
    fn test_import_json_array_of_arrays() {
        let temp_dir = TempDir::new().unwrap();
        let json_path = temp_dir.path().join("data.json");
        let xlsx_path = temp_dir.path().join("output.xlsx");

        std::fs::write(&json_path, r#"[["Name", "Age"], ["Alice", 30]]"#).unwrap();

        let result = import_json(&json_path, &xlsx_path, None, &default_global());
        assert!(result.is_ok());
    }

    #[test]
    fn test_import_ndjson() {
        let temp_dir = TempDir::new().unwrap();
        let ndjson_path = temp_dir.path().join("data.ndjson");
        let xlsx_path = temp_dir.path().join("output.xlsx");

        std::fs::write(
            &ndjson_path,
            r#"{"name": "Alice", "age": 30}
{"name": "Bob", "age": 25}"#,
        )
        .unwrap();

        let result = import_ndjson(&ndjson_path, &xlsx_path, None, true, &default_global());
        assert!(result.is_ok());
        assert!(xlsx_path.exists());
    }

    #[test]
    fn test_import_ndjson_array_format() {
        let temp_dir = TempDir::new().unwrap();
        let ndjson_path = temp_dir.path().join("data.ndjson");
        let xlsx_path = temp_dir.path().join("output.xlsx");

        std::fs::write(
            &ndjson_path,
            r#"["Alice", 30]
["Bob", 25]"#,
        )
        .unwrap();

        let result = import_ndjson(&ndjson_path, &xlsx_path, None, false, &default_global());
        assert!(result.is_ok());
    }

    // Additional tests for better coverage

    #[test]
    fn test_import_csv_with_sheet() {
        let temp_dir = TempDir::new().unwrap();
        let csv_path = temp_dir.path().join("data_sheet.csv");
        let xlsx_path = temp_dir.path().join("output_sheet.xlsx");

        std::fs::write(&csv_path, "A,B\n1,2\n3,4").unwrap();

        let result = import_csv(
            &csv_path,
            &xlsx_path,
            Some("Data"),
            ',',
            false,
            &default_global(),
        );
        assert!(result.is_ok());

        let wb = Workbook::open(&xlsx_path).unwrap();
        assert!(wb.get_sheet("Data").is_some());
    }

    #[test]
    fn test_import_csv_with_header() {
        let temp_dir = TempDir::new().unwrap();
        let csv_path = temp_dir.path().join("data_header.csv");
        let xlsx_path = temp_dir.path().join("output_header.xlsx");

        std::fs::write(&csv_path, "Name,Age\nAlice,30\nBob,25").unwrap();

        let result = import_csv(&csv_path, &xlsx_path, None, ',', true, &default_global());
        assert!(result.is_ok());
    }

    #[test]
    fn test_import_csv_semicolon_delimiter() {
        let temp_dir = TempDir::new().unwrap();
        let csv_path = temp_dir.path().join("data_semi.csv");
        let xlsx_path = temp_dir.path().join("output_semi.xlsx");

        std::fs::write(&csv_path, "Name;Age\nAlice;30").unwrap();

        let result = import_csv(&csv_path, &xlsx_path, None, ';', false, &default_global());
        assert!(result.is_ok());
    }

    #[test]
    fn test_import_json_dry_run() {
        let temp_dir = TempDir::new().unwrap();
        let json_path = temp_dir.path().join("data_dry.json");
        let xlsx_path = temp_dir.path().join("output_dry.xlsx");

        std::fs::write(&json_path, r#"[{"a": 1}]"#).unwrap();

        let mut global = default_global();
        global.dry_run = true;

        let result = import_json(&json_path, &xlsx_path, None, &global);
        assert!(result.is_ok());
        assert!(!xlsx_path.exists());
    }

    #[test]
    fn test_import_ndjson_dry_run() {
        let temp_dir = TempDir::new().unwrap();
        let ndjson_path = temp_dir.path().join("data_dry.ndjson");
        let xlsx_path = temp_dir.path().join("output_dry.xlsx");

        std::fs::write(&ndjson_path, r#"{"a": 1}"#).unwrap();

        let mut global = default_global();
        global.dry_run = true;

        let result = import_ndjson(&ndjson_path, &xlsx_path, None, true, &global);
        assert!(result.is_ok());
        assert!(!xlsx_path.exists());
    }

    #[test]
    fn test_import_tsv_with_sheet() {
        let temp_dir = TempDir::new().unwrap();
        let tsv_path = temp_dir.path().join("data_sheet.tsv");
        let xlsx_path = temp_dir.path().join("output_sheet.xlsx");

        std::fs::write(&tsv_path, "A\tB\n1\t2").unwrap();

        let result = import_tsv(&tsv_path, &xlsx_path, Some("MySheet"), &default_global());
        assert!(result.is_ok());

        let wb = Workbook::open(&xlsx_path).unwrap();
        assert!(wb.get_sheet("MySheet").is_some());
    }

    #[test]
    fn test_import_json_with_sheet() {
        let temp_dir = TempDir::new().unwrap();
        let json_path = temp_dir.path().join("data_sheet.json");
        let xlsx_path = temp_dir.path().join("output_sheet.xlsx");

        std::fs::write(&json_path, r#"[{"x": 1}]"#).unwrap();

        let result = import_json(&json_path, &xlsx_path, Some("JsonData"), &default_global());
        assert!(result.is_ok());

        let wb = Workbook::open(&xlsx_path).unwrap();
        assert!(wb.get_sheet("JsonData").is_some());
    }

    #[test]
    fn test_import_ndjson_with_sheet() {
        let temp_dir = TempDir::new().unwrap();
        let ndjson_path = temp_dir.path().join("data_sheet.ndjson");
        let xlsx_path = temp_dir.path().join("output_sheet.xlsx");

        std::fs::write(&ndjson_path, r#"{"y": 2}"#).unwrap();

        let result = import_ndjson(
            &ndjson_path,
            &xlsx_path,
            Some("NdjsonData"),
            true,
            &default_global(),
        );
        assert!(result.is_ok());

        let wb = Workbook::open(&xlsx_path).unwrap();
        assert!(wb.get_sheet("NdjsonData").is_some());
    }

    #[test]
    fn test_run_csv_command() {
        let temp_dir = TempDir::new().unwrap();
        let csv_path = temp_dir.path().join("run.csv");
        let xlsx_path = temp_dir.path().join("run.xlsx");

        std::fs::write(&csv_path, "A,B\n1,2").unwrap();

        let args = ImportArgs {
            command: ImportCommand::Csv {
                source: csv_path,
                dest: xlsx_path.clone(),
                sheet: None,
                delimiter: ',',
                header: false,
            },
        };

        let result = run(&args, &default_global());
        assert!(result.is_ok());
    }

    #[test]
    fn test_run_tsv_command() {
        let temp_dir = TempDir::new().unwrap();
        let tsv_path = temp_dir.path().join("run.tsv");
        let xlsx_path = temp_dir.path().join("run.xlsx");

        std::fs::write(&tsv_path, "A\tB\n1\t2").unwrap();

        let args = ImportArgs {
            command: ImportCommand::Tsv {
                source: tsv_path,
                dest: xlsx_path,
                sheet: None,
            },
        };

        let result = run(&args, &default_global());
        assert!(result.is_ok());
    }

    #[test]
    fn test_run_json_command() {
        let temp_dir = TempDir::new().unwrap();
        let json_path = temp_dir.path().join("run.json");
        let xlsx_path = temp_dir.path().join("run.xlsx");

        std::fs::write(&json_path, r#"[{"a": 1}]"#).unwrap();

        let args = ImportArgs {
            command: ImportCommand::Json {
                source: json_path,
                dest: xlsx_path,
                sheet: None,
            },
        };

        let result = run(&args, &default_global());
        assert!(result.is_ok());
    }

    #[test]
    fn test_run_ndjson_command() {
        let temp_dir = TempDir::new().unwrap();
        let ndjson_path = temp_dir.path().join("run.ndjson");
        let xlsx_path = temp_dir.path().join("run.xlsx");

        std::fs::write(&ndjson_path, r#"{"b": 2}"#).unwrap();

        let args = ImportArgs {
            command: ImportCommand::Ndjson {
                source: ndjson_path,
                dest: xlsx_path,
                sheet: None,
                header: true,
            },
        };

        let result = run(&args, &default_global());
        assert!(result.is_ok());
    }

    #[test]
    fn test_json_to_cell_array() {
        let arr = serde_json::json!([1, 2, 3]);
        let result = json_to_cell(&arr);
        // Arrays are converted to strings
        assert!(matches!(result, CellValue::String(_)));
    }

    #[test]
    fn test_json_to_cell_object() {
        let obj = serde_json::json!({"key": "value"});
        let result = json_to_cell(&obj);
        // Objects are converted to strings
        assert!(matches!(result, CellValue::String(_)));
    }

    #[test]
    fn test_parse_value_yes_no() {
        assert_eq!(parse_value("YES"), CellValue::Boolean(true));
        assert_eq!(parse_value("no"), CellValue::Boolean(false));
        assert_eq!(parse_value("NO"), CellValue::Boolean(false));
    }
}
