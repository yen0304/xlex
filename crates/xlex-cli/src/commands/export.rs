//! Export operations.

use anyhow::Result;
use clap::{Parser, Subcommand};
use colored::Colorize;

use xlex_core::Workbook;

use super::GlobalOptions;

/// Arguments for export operations.
#[derive(Parser)]
pub struct ExportArgs {
    #[command(subcommand)]
    pub command: ExportCommand,
}

#[derive(Subcommand)]
pub enum ExportCommand {
    /// Export to CSV
    Csv {
        /// Source xlsx file
        source: std::path::PathBuf,
        /// Destination file (or - for stdout)
        dest: String,
        /// Sheet name (default: first sheet)
        #[arg(short, long)]
        sheet: Option<String>,
        /// Delimiter character
        #[arg(short, long, default_value = ",")]
        delimiter: char,
        /// Export all sheets (creates multiple files with sheet name suffix)
        #[arg(long)]
        all: bool,
    },
    /// Export to JSON
    Json {
        /// Source xlsx file
        source: std::path::PathBuf,
        /// Destination file (or - for stdout)
        dest: String,
        /// Sheet name (default: first sheet)
        #[arg(short, long)]
        sheet: Option<String>,
        /// Use first row as keys
        #[arg(long)]
        header: bool,
        /// Export all sheets (creates multiple files or combined JSON)
        #[arg(long)]
        all: bool,
    },
    /// Export to TSV
    Tsv {
        /// Source xlsx file
        source: std::path::PathBuf,
        /// Destination file (or - for stdout)
        dest: String,
        /// Sheet name (default: first sheet)
        #[arg(short, long)]
        sheet: Option<String>,
        /// Export all sheets
        #[arg(long)]
        all: bool,
    },
    /// Export to YAML
    Yaml {
        /// Source xlsx file
        source: std::path::PathBuf,
        /// Destination file (or - for stdout)
        dest: String,
        /// Sheet name (default: first sheet)
        #[arg(short, long)]
        sheet: Option<String>,
        /// Export all sheets
        #[arg(long)]
        all: bool,
    },
    /// Export to Markdown table
    Markdown {
        /// Source xlsx file
        source: std::path::PathBuf,
        /// Destination file (or - for stdout)
        dest: String,
        /// Sheet name (default: first sheet)
        #[arg(short, long)]
        sheet: Option<String>,
        /// Export all sheets
        #[arg(long)]
        all: bool,
    },
    /// Export to NDJSON (newline-delimited JSON)
    Ndjson {
        /// Source xlsx file
        source: std::path::PathBuf,
        /// Destination file (or - for stdout)
        dest: String,
        /// Sheet name (default: first sheet)
        #[arg(short, long)]
        sheet: Option<String>,
        /// Use first row as keys
        #[arg(long)]
        header: bool,
        /// Export all sheets
        #[arg(long)]
        all: bool,
    },
    /// Export workbook metadata
    Meta {
        /// Source xlsx file
        source: std::path::PathBuf,
        /// Destination file (or - for stdout)
        dest: String,
    },
}

/// Run export operations.
pub fn run(args: &ExportArgs, global: &GlobalOptions) -> Result<()> {
    match &args.command {
        ExportCommand::Csv {
            source,
            dest,
            sheet,
            delimiter,
            all,
        } => {
            if *all {
                export_all_csv(source, dest, *delimiter, global)
            } else {
                export_csv(source, dest, sheet.as_deref(), *delimiter, global)
            }
        }
        ExportCommand::Json {
            source,
            dest,
            sheet,
            header,
            all,
        } => {
            if *all {
                export_all_json(source, dest, *header, global)
            } else {
                export_json(source, dest, sheet.as_deref(), *header, global)
            }
        }
        ExportCommand::Tsv {
            source,
            dest,
            sheet,
            all,
        } => {
            if *all {
                export_all_csv(source, dest, '\t', global) // TSV is CSV with tab
            } else {
                export_tsv(source, dest, sheet.as_deref(), global)
            }
        }
        ExportCommand::Yaml {
            source,
            dest,
            sheet,
            all,
        } => {
            if *all {
                export_all_yaml(source, dest, global)
            } else {
                export_yaml(source, dest, sheet.as_deref(), global)
            }
        }
        ExportCommand::Markdown {
            source,
            dest,
            sheet,
            all,
        } => {
            if *all {
                export_all_markdown(source, dest, global)
            } else {
                export_markdown(source, dest, sheet.as_deref(), global)
            }
        }
        ExportCommand::Ndjson {
            source,
            dest,
            sheet,
            header,
            all,
        } => {
            if *all {
                export_all_ndjson(source, dest, *header, global)
            } else {
                export_ndjson(source, dest, sheet.as_deref(), *header, global)
            }
        }
        ExportCommand::Meta { source, dest } => export_meta(source, dest, global),
    }
}

fn export_all_csv(
    source: &std::path::Path,
    dest: &str,
    delimiter: char,
    global: &GlobalOptions,
) -> Result<()> {
    let workbook = Workbook::open(source)?;
    let sheet_names: Vec<String> = workbook.sheet_names().iter().map(|s| s.to_string()).collect();

    // Create output directory based on dest name
    let base_path = std::path::Path::new(dest);
    let parent = base_path.parent().unwrap_or(std::path::Path::new("."));
    let stem = base_path.file_stem().map(|s| s.to_string_lossy().to_string()).unwrap_or_else(|| "export".to_string());
    let ext = base_path.extension().map(|s| s.to_string_lossy().to_string()).unwrap_or_else(|| "csv".to_string());

    for sheet_name in &sheet_names {
        let output_path = parent.join(format!("{}_{}.{}", stem, sheet_name.replace(' ', "_"), ext));
        export_csv(source, &output_path.to_string_lossy(), Some(sheet_name), delimiter, global)?;
    }

    if !global.quiet {
        println!("{} Exported {} sheets", "✓".green(), sheet_names.len());
    }

    Ok(())
}

fn export_all_json(
    source: &std::path::Path,
    dest: &str,
    header: bool,
    global: &GlobalOptions,
) -> Result<()> {
    let workbook = Workbook::open(source)?;
    let sheet_names: Vec<String> = workbook.sheet_names().iter().map(|s| s.to_string()).collect();

    // Combined JSON with all sheets
    let mut combined: serde_json::Map<String, serde_json::Value> = serde_json::Map::new();

    for sheet_name in &sheet_names {
        let sheet_obj = workbook
            .get_sheet(sheet_name)
            .ok_or_else(|| xlex_core::XlexError::SheetNotFound {
                name: sheet_name.to_string(),
            })?;

        let (max_col, max_row) = sheet_obj.dimensions();
        let data = export_sheet_to_json_value(sheet_obj, max_col, max_row, header);
        combined.insert(sheet_name.clone(), data);
    }

    let json_str = serde_json::to_string_pretty(&serde_json::Value::Object(combined))?;
    write_output(dest, &json_str, global)?;

    if !global.quiet && dest != "-" {
        println!("{} Exported {} sheets to {}", "✓".green(), sheet_names.len(), dest);
    }

    Ok(())
}

fn export_all_yaml(
    source: &std::path::Path,
    dest: &str,
    global: &GlobalOptions,
) -> Result<()> {
    let workbook = Workbook::open(source)?;
    let sheet_names: Vec<String> = workbook.sheet_names().iter().map(|s| s.to_string()).collect();

    let base_path = std::path::Path::new(dest);
    let parent = base_path.parent().unwrap_or(std::path::Path::new("."));
    let stem = base_path.file_stem().map(|s| s.to_string_lossy().to_string()).unwrap_or_else(|| "export".to_string());
    let ext = base_path.extension().map(|s| s.to_string_lossy().to_string()).unwrap_or_else(|| "yaml".to_string());

    for sheet_name in &sheet_names {
        let output_path = parent.join(format!("{}_{}.{}", stem, sheet_name.replace(' ', "_"), ext));
        export_yaml(source, &output_path.to_string_lossy(), Some(sheet_name), global)?;
    }

    if !global.quiet {
        println!("{} Exported {} sheets", "✓".green(), sheet_names.len());
    }

    Ok(())
}

fn export_all_markdown(
    source: &std::path::Path,
    dest: &str,
    global: &GlobalOptions,
) -> Result<()> {
    let workbook = Workbook::open(source)?;
    let sheet_names: Vec<String> = workbook.sheet_names().iter().map(|s| s.to_string()).collect();

    let base_path = std::path::Path::new(dest);
    let parent = base_path.parent().unwrap_or(std::path::Path::new("."));
    let stem = base_path.file_stem().map(|s| s.to_string_lossy().to_string()).unwrap_or_else(|| "export".to_string());
    let ext = base_path.extension().map(|s| s.to_string_lossy().to_string()).unwrap_or_else(|| "md".to_string());

    for sheet_name in &sheet_names {
        let output_path = parent.join(format!("{}_{}.{}", stem, sheet_name.replace(' ', "_"), ext));
        export_markdown(source, &output_path.to_string_lossy(), Some(sheet_name), global)?;
    }

    if !global.quiet {
        println!("{} Exported {} sheets", "✓".green(), sheet_names.len());
    }

    Ok(())
}

fn export_all_ndjson(
    source: &std::path::Path,
    dest: &str,
    header: bool,
    global: &GlobalOptions,
) -> Result<()> {
    let workbook = Workbook::open(source)?;
    let sheet_names: Vec<String> = workbook.sheet_names().iter().map(|s| s.to_string()).collect();

    let base_path = std::path::Path::new(dest);
    let parent = base_path.parent().unwrap_or(std::path::Path::new("."));
    let stem = base_path.file_stem().map(|s| s.to_string_lossy().to_string()).unwrap_or_else(|| "export".to_string());
    let ext = base_path.extension().map(|s| s.to_string_lossy().to_string()).unwrap_or_else(|| "ndjson".to_string());

    for sheet_name in &sheet_names {
        let output_path = parent.join(format!("{}_{}.{}", stem, sheet_name.replace(' ', "_"), ext));
        export_ndjson(source, &output_path.to_string_lossy(), Some(sheet_name), header, global)?;
    }

    if !global.quiet {
        println!("{} Exported {} sheets", "✓".green(), sheet_names.len());
    }

    Ok(())
}

fn export_sheet_to_json_value(
    sheet: &xlex_core::Sheet,
    max_col: u32,
    max_row: u32,
    header: bool,
) -> serde_json::Value {
    use xlex_core::CellValue;

    if header && max_row > 0 {
        // Export as array of objects
        let mut headers: Vec<String> = Vec::new();
        for col in 1..=max_col {
            let cell_ref = xlex_core::CellRef::new(col, 1);
            let value = sheet.get_value(&cell_ref);
            headers.push(value.to_display_string());
        }

        let mut rows = Vec::new();
        for row in 2..=max_row {
            let mut obj = serde_json::Map::new();
            for (idx, col) in (1..=max_col).enumerate() {
                let cell_ref = xlex_core::CellRef::new(col, row);
                let value = sheet.get_value(&cell_ref);
                let json_value = match value {
                    CellValue::Empty => serde_json::Value::Null,
                    CellValue::String(s) => serde_json::Value::String(s),
                    CellValue::Number(n) => serde_json::json!(n),
                    CellValue::Boolean(b) => serde_json::Value::Bool(b),
                    _ => serde_json::Value::String(value.to_display_string()),
                };
                if idx < headers.len() {
                    obj.insert(headers[idx].clone(), json_value);
                }
            }
            rows.push(serde_json::Value::Object(obj));
        }
        serde_json::Value::Array(rows)
    } else {
        // Export as 2D array
        let mut rows = Vec::new();
        for row in 1..=max_row {
            let mut row_values = Vec::new();
            for col in 1..=max_col {
                let cell_ref = xlex_core::CellRef::new(col, row);
                let value = sheet.get_value(&cell_ref);
                let json_value = match value {
                    CellValue::Empty => serde_json::Value::Null,
                    CellValue::String(s) => serde_json::Value::String(s),
                    CellValue::Number(n) => serde_json::json!(n),
                    CellValue::Boolean(b) => serde_json::Value::Bool(b),
                    _ => serde_json::Value::String(value.to_display_string()),
                };
                row_values.push(json_value);
            }
            rows.push(serde_json::Value::Array(row_values));
        }
        serde_json::Value::Array(rows)
    }
}

fn export_csv(
    source: &std::path::Path,
    dest: &str,
    sheet: Option<&str>,
    delimiter: char,
    global: &GlobalOptions,
) -> Result<()> {
    let workbook = Workbook::open(source)?;
    let sheet_name = sheet
        .or_else(|| workbook.sheet_names().first().copied())
        .ok_or_else(|| anyhow::anyhow!("No sheets in workbook"))?;

    let sheet_obj = workbook
        .get_sheet(sheet_name)
        .ok_or_else(|| xlex_core::XlexError::SheetNotFound {
            name: sheet_name.to_string(),
        })?;

    // Get dimensions
    let (max_col, max_row) = sheet_obj.dimensions();

    let mut output = String::new();
    for row in 1..=max_row {
        let mut row_values: Vec<String> = Vec::new();
        for col in 1..=max_col {
            let cell_ref = xlex_core::CellRef::new(col, row);
            let value = sheet_obj.get_value(&cell_ref);
            let str_value = value.to_display_string();
            // Escape if contains delimiter or newline
            if str_value.contains(delimiter) || str_value.contains('\n') || str_value.contains('"')
            {
                row_values.push(format!("\"{}\"", str_value.replace('"', "\"\"")));
            } else {
                row_values.push(str_value);
            }
        }
        output.push_str(&row_values.join(&delimiter.to_string()));
        output.push('\n');
    }

    write_output(dest, &output, global)?;

    if !global.quiet && dest != "-" {
        println!("Exported {} rows to {}", max_row.to_string().green(), dest);
    }

    Ok(())
}

fn export_json(
    source: &std::path::Path,
    dest: &str,
    sheet: Option<&str>,
    has_header: bool,
    global: &GlobalOptions,
) -> Result<()> {
    let workbook = Workbook::open(source)?;
    let sheet_name = sheet
        .or_else(|| workbook.sheet_names().first().copied())
        .ok_or_else(|| anyhow::anyhow!("No sheets in workbook"))?;

    let sheet_obj = workbook
        .get_sheet(sheet_name)
        .ok_or_else(|| xlex_core::XlexError::SheetNotFound {
            name: sheet_name.to_string(),
        })?;

    // Get dimensions
    let (max_col, max_row) = sheet_obj.dimensions();

    let data = if has_header && max_row > 0 {
        // Use first row as keys
        let mut headers: Vec<String> = Vec::new();
        for col in 1..=max_col {
            let cell_ref = xlex_core::CellRef::new(col, 1);
            let value = sheet_obj.get_value(&cell_ref);
            headers.push(value.to_display_string());
        }

        let mut rows: Vec<serde_json::Value> = Vec::new();
        for row in 2..=max_row {
            let mut obj = serde_json::Map::new();
            for (col_idx, header) in headers.iter().enumerate() {
                let cell_ref = xlex_core::CellRef::new((col_idx + 1) as u32, row);
                let value = sheet_obj.get_value(&cell_ref);
                obj.insert(header.clone(), cell_to_json(&value));
            }
            rows.push(serde_json::Value::Object(obj));
        }
        serde_json::Value::Array(rows)
    } else {
        // Array of arrays
        let mut rows: Vec<serde_json::Value> = Vec::new();
        for row in 1..=max_row {
            let mut row_values: Vec<serde_json::Value> = Vec::new();
            for col in 1..=max_col {
                let cell_ref = xlex_core::CellRef::new(col, row);
                let value = sheet_obj.get_value(&cell_ref);
                row_values.push(cell_to_json(&value));
            }
            rows.push(serde_json::Value::Array(row_values));
        }
        serde_json::Value::Array(rows)
    };

    let output = serde_json::to_string_pretty(&data)?;
    write_output(dest, &output, global)?;

    if !global.quiet && dest != "-" {
        println!("Exported JSON to {}", dest);
    }

    Ok(())
}

fn export_tsv(
    source: &std::path::Path,
    dest: &str,
    sheet: Option<&str>,
    global: &GlobalOptions,
) -> Result<()> {
    export_csv(source, dest, sheet, '\t', global)
}

fn export_yaml(
    source: &std::path::Path,
    dest: &str,
    sheet: Option<&str>,
    global: &GlobalOptions,
) -> Result<()> {
    let workbook = Workbook::open(source)?;
    let sheet_name = sheet
        .or_else(|| workbook.sheet_names().first().copied())
        .ok_or_else(|| anyhow::anyhow!("No sheets in workbook"))?;

    let sheet_obj = workbook
        .get_sheet(sheet_name)
        .ok_or_else(|| xlex_core::XlexError::SheetNotFound {
            name: sheet_name.to_string(),
        })?;

    // Get dimensions
    let (max_col, max_row) = sheet_obj.dimensions();

    let mut rows: Vec<Vec<serde_json::Value>> = Vec::new();
    for row in 1..=max_row {
        let mut row_values: Vec<serde_json::Value> = Vec::new();
        for col in 1..=max_col {
            let cell_ref = xlex_core::CellRef::new(col, row);
            let value = sheet_obj.get_value(&cell_ref);
            row_values.push(cell_to_json(&value));
        }
        rows.push(row_values);
    }

    let output = serde_yaml::to_string(&rows)?;
    write_output(dest, &output, global)?;

    if !global.quiet && dest != "-" {
        println!("Exported YAML to {}", dest);
    }

    Ok(())
}

fn export_markdown(
    source: &std::path::Path,
    dest: &str,
    sheet: Option<&str>,
    global: &GlobalOptions,
) -> Result<()> {
    let workbook = Workbook::open(source)?;
    let sheet_name = sheet
        .or_else(|| workbook.sheet_names().first().copied())
        .ok_or_else(|| anyhow::anyhow!("No sheets in workbook"))?;

    let sheet_obj = workbook
        .get_sheet(sheet_name)
        .ok_or_else(|| xlex_core::XlexError::SheetNotFound {
            name: sheet_name.to_string(),
        })?;

    // Get dimensions
    let (max_col, max_row) = sheet_obj.dimensions();

    let mut output = String::new();

    // First row is header
    if max_row > 0 {
        let mut header_row: Vec<String> = Vec::new();
        for col in 1..=max_col {
            let cell_ref = xlex_core::CellRef::new(col, 1);
            let value = sheet_obj.get_value(&cell_ref);
            header_row.push(value.to_display_string());
        }
        output.push_str(&format!("| {} |\n", header_row.join(" | ")));

        // Separator
        let separator: Vec<&str> = vec!["---"; max_col as usize];
        output.push_str(&format!("| {} |\n", separator.join(" | ")));

        // Data rows
        for row in 2..=max_row {
            let mut row_values: Vec<String> = Vec::new();
            for col in 1..=max_col {
                let cell_ref = xlex_core::CellRef::new(col, row);
                let value = sheet_obj.get_value(&cell_ref);
                row_values.push(value.to_display_string());
            }
            output.push_str(&format!("| {} |\n", row_values.join(" | ")));
        }
    }

    write_output(dest, &output, global)?;

    if !global.quiet && dest != "-" {
        println!("Exported Markdown to {}", dest);
    }

    Ok(())
}

fn cell_to_json(value: &xlex_core::CellValue) -> serde_json::Value {
    match value {
        xlex_core::CellValue::Empty => serde_json::Value::Null,
        xlex_core::CellValue::String(s) => serde_json::Value::String(s.clone()),
        xlex_core::CellValue::Number(n) => serde_json::json!(*n),
        xlex_core::CellValue::Boolean(b) => serde_json::Value::Bool(*b),
        _ => serde_json::Value::String(value.to_display_string()),
    }
}

fn write_output(dest: &str, content: &str, _global: &GlobalOptions) -> Result<()> {
    if dest == "-" {
        print!("{}", content);
    } else {
        std::fs::write(dest, content)?;
    }
    Ok(())
}

fn export_ndjson(
    source: &std::path::Path,
    dest: &str,
    sheet: Option<&str>,
    has_header: bool,
    global: &GlobalOptions,
) -> Result<()> {
    let workbook = Workbook::open(source)?;
    let sheet_name = sheet
        .or_else(|| workbook.sheet_names().first().copied())
        .ok_or_else(|| anyhow::anyhow!("No sheets in workbook"))?;

    let sheet_obj = workbook
        .get_sheet(sheet_name)
        .ok_or_else(|| xlex_core::XlexError::SheetNotFound {
            name: sheet_name.to_string(),
        })?;

    // Get dimensions
    let (max_col, max_row) = sheet_obj.dimensions();

    let mut output = String::new();

    if has_header && max_row > 0 {
        // Use first row as keys
        let mut headers: Vec<String> = Vec::new();
        for col in 1..=max_col {
            let cell_ref = xlex_core::CellRef::new(col, 1);
            let value = sheet_obj.get_value(&cell_ref);
            headers.push(value.to_display_string());
        }

        for row in 2..=max_row {
            let mut obj = serde_json::Map::new();
            for (col_idx, header) in headers.iter().enumerate() {
                let cell_ref = xlex_core::CellRef::new((col_idx + 1) as u32, row);
                let value = sheet_obj.get_value(&cell_ref);
                obj.insert(header.clone(), cell_to_json(&value));
            }
            output.push_str(&serde_json::to_string(&serde_json::Value::Object(obj))?);
            output.push('\n');
        }
    } else {
        // Array of arrays, one per line
        for row in 1..=max_row {
            let mut row_values: Vec<serde_json::Value> = Vec::new();
            for col in 1..=max_col {
                let cell_ref = xlex_core::CellRef::new(col, row);
                let value = sheet_obj.get_value(&cell_ref);
                row_values.push(cell_to_json(&value));
            }
            output.push_str(&serde_json::to_string(&serde_json::Value::Array(row_values))?);
            output.push('\n');
        }
    }

    write_output(dest, &output, global)?;

    if !global.quiet && dest != "-" {
        println!("Exported {} rows to NDJSON {}", max_row.to_string().green(), dest);
    }

    Ok(())
}

fn export_meta(
    source: &std::path::Path,
    dest: &str,
    global: &GlobalOptions,
) -> Result<()> {
    let workbook = Workbook::open(source)?;

    // Build metadata
    let mut sheets_meta: Vec<serde_json::Value> = Vec::new();
    
    for name in workbook.sheet_names() {
        if let Some(sheet) = workbook.get_sheet(name) {
            let (max_col, max_row) = sheet.dimensions();
            let cell_count = sheet.cell_count();
            let merged_ranges = sheet.merged_ranges();
            
            sheets_meta.push(serde_json::json!({
                "name": name,
                "index": sheet.info.sheet_id,
                "visibility": sheet.info.visibility.to_string(),
                "dimensions": {
                    "columns": max_col,
                    "rows": max_row,
                },
                "cellCount": cell_count,
                "mergedRanges": merged_ranges.len(),
            }));
        }
    }

    let meta = serde_json::json!({
        "file": source.file_name().map(|n| n.to_string_lossy()).unwrap_or_default(),
        "path": source.to_string_lossy(),
        "sheets": {
            "count": workbook.sheet_count(),
            "names": workbook.sheet_names(),
        },
        "sheetDetails": sheets_meta,
    });

    let output = serde_json::to_string_pretty(&meta)?;
    write_output(dest, &output, global)?;

    if !global.quiet && dest != "-" {
        println!("Exported metadata to {}", dest);
    }

    Ok(())
}
