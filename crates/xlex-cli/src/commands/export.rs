//! Export operations.

use anyhow::Result;
use clap::{Parser, Subcommand};
use colored::Colorize;

use xlex_core::Workbook;

use super::GlobalOptions;
use crate::progress::Progress;

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
    let sheet_names: Vec<String> = workbook
        .sheet_names()
        .iter()
        .map(|s| s.to_string())
        .collect();

    // Create output directory based on dest name
    let base_path = std::path::Path::new(dest);
    let parent = base_path.parent().unwrap_or(std::path::Path::new("."));
    let stem = base_path
        .file_stem()
        .map(|s| s.to_string_lossy().to_string())
        .unwrap_or_else(|| "export".to_string());
    let ext = base_path
        .extension()
        .map(|s| s.to_string_lossy().to_string())
        .unwrap_or_else(|| "csv".to_string());

    for sheet_name in &sheet_names {
        let output_path = parent.join(format!("{}_{}.{}", stem, sheet_name.replace(' ', "_"), ext));
        export_csv(
            source,
            &output_path.to_string_lossy(),
            Some(sheet_name),
            delimiter,
            global,
        )?;
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
    let sheet_names: Vec<String> = workbook
        .sheet_names()
        .iter()
        .map(|s| s.to_string())
        .collect();

    // Combined JSON with all sheets
    let mut combined: serde_json::Map<String, serde_json::Value> = serde_json::Map::new();

    for sheet_name in &sheet_names {
        let sheet_obj =
            workbook
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
        println!(
            "{} Exported {} sheets to {}",
            "✓".green(),
            sheet_names.len(),
            dest
        );
    }

    Ok(())
}

fn export_all_yaml(source: &std::path::Path, dest: &str, global: &GlobalOptions) -> Result<()> {
    let workbook = Workbook::open(source)?;
    let sheet_names: Vec<String> = workbook
        .sheet_names()
        .iter()
        .map(|s| s.to_string())
        .collect();

    let base_path = std::path::Path::new(dest);
    let parent = base_path.parent().unwrap_or(std::path::Path::new("."));
    let stem = base_path
        .file_stem()
        .map(|s| s.to_string_lossy().to_string())
        .unwrap_or_else(|| "export".to_string());
    let ext = base_path
        .extension()
        .map(|s| s.to_string_lossy().to_string())
        .unwrap_or_else(|| "yaml".to_string());

    for sheet_name in &sheet_names {
        let output_path = parent.join(format!("{}_{}.{}", stem, sheet_name.replace(' ', "_"), ext));
        export_yaml(
            source,
            &output_path.to_string_lossy(),
            Some(sheet_name),
            global,
        )?;
    }

    if !global.quiet {
        println!("{} Exported {} sheets", "✓".green(), sheet_names.len());
    }

    Ok(())
}

fn export_all_markdown(source: &std::path::Path, dest: &str, global: &GlobalOptions) -> Result<()> {
    let workbook = Workbook::open(source)?;
    let sheet_names: Vec<String> = workbook
        .sheet_names()
        .iter()
        .map(|s| s.to_string())
        .collect();

    let base_path = std::path::Path::new(dest);
    let parent = base_path.parent().unwrap_or(std::path::Path::new("."));
    let stem = base_path
        .file_stem()
        .map(|s| s.to_string_lossy().to_string())
        .unwrap_or_else(|| "export".to_string());
    let ext = base_path
        .extension()
        .map(|s| s.to_string_lossy().to_string())
        .unwrap_or_else(|| "md".to_string());

    for sheet_name in &sheet_names {
        let output_path = parent.join(format!("{}_{}.{}", stem, sheet_name.replace(' ', "_"), ext));
        export_markdown(
            source,
            &output_path.to_string_lossy(),
            Some(sheet_name),
            global,
        )?;
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
    let sheet_names: Vec<String> = workbook
        .sheet_names()
        .iter()
        .map(|s| s.to_string())
        .collect();

    let base_path = std::path::Path::new(dest);
    let parent = base_path.parent().unwrap_or(std::path::Path::new("."));
    let stem = base_path
        .file_stem()
        .map(|s| s.to_string_lossy().to_string())
        .unwrap_or_else(|| "export".to_string());
    let ext = base_path
        .extension()
        .map(|s| s.to_string_lossy().to_string())
        .unwrap_or_else(|| "ndjson".to_string());

    for sheet_name in &sheet_names {
        let output_path = parent.join(format!("{}_{}.{}", stem, sheet_name.replace(' ', "_"), ext));
        export_ndjson(
            source,
            &output_path.to_string_lossy(),
            Some(sheet_name),
            header,
            global,
        )?;
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

    let sheet_obj =
        workbook
            .get_sheet(sheet_name)
            .ok_or_else(|| xlex_core::XlexError::SheetNotFound {
                name: sheet_name.to_string(),
            })?;

    // Get dimensions
    let (max_col, max_row) = sheet_obj.dimensions();

    // Create progress for large exports
    let progress = if max_row > 100 {
        Some(Progress::bar(
            max_row as u64,
            "Exporting to CSV...",
            global.quiet,
        ))
    } else {
        None
    };

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

        if let Some(ref pb) = progress {
            pb.inc(1);
        }
    }

    if let Some(ref pb) = progress {
        pb.finish_and_clear();
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

    let sheet_obj =
        workbook
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

    let sheet_obj =
        workbook
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

    let sheet_obj =
        workbook
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

    let sheet_obj =
        workbook
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
            output.push_str(&serde_json::to_string(&serde_json::Value::Array(
                row_values,
            ))?);
            output.push('\n');
        }
    }

    write_output(dest, &output, global)?;

    if !global.quiet && dest != "-" {
        println!(
            "Exported {} rows to NDJSON {}",
            max_row.to_string().green(),
            dest
        );
    }

    Ok(())
}

fn export_meta(source: &std::path::Path, dest: &str, global: &GlobalOptions) -> Result<()> {
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

#[cfg(test)]
mod tests {
    use super::*;
    use tempfile::TempDir;
    use xlex_core::CellValue;

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

    fn create_test_workbook(dir: &TempDir, name: &str) -> std::path::PathBuf {
        let file_path = dir.path().join(name);
        let wb = Workbook::new();
        wb.save_as(&file_path).unwrap();
        file_path
    }

    fn setup_test_data(file: &std::path::Path) {
        let mut wb = Workbook::open(file).unwrap();
        // Header row
        wb.set_cell(
            "Sheet1",
            xlex_core::CellRef::new(1, 1),
            CellValue::String("Name".to_string()),
        )
        .unwrap();
        wb.set_cell(
            "Sheet1",
            xlex_core::CellRef::new(2, 1),
            CellValue::String("Age".to_string()),
        )
        .unwrap();
        // Data rows
        wb.set_cell(
            "Sheet1",
            xlex_core::CellRef::new(1, 2),
            CellValue::String("Alice".to_string()),
        )
        .unwrap();
        wb.set_cell(
            "Sheet1",
            xlex_core::CellRef::new(2, 2),
            CellValue::Number(30.0),
        )
        .unwrap();
        wb.set_cell(
            "Sheet1",
            xlex_core::CellRef::new(1, 3),
            CellValue::String("Bob".to_string()),
        )
        .unwrap();
        wb.set_cell(
            "Sheet1",
            xlex_core::CellRef::new(2, 3),
            CellValue::Number(25.0),
        )
        .unwrap();
        wb.save().unwrap();
    }

    #[test]
    fn test_export_csv() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "test.xlsx");
        setup_test_data(&file_path);

        let dest = temp_dir.path().join("output.csv");
        let result = export_csv(
            &file_path,
            &dest.to_string_lossy(),
            None,
            ',',
            &default_global(),
        );
        assert!(result.is_ok());
        assert!(dest.exists());

        let content = std::fs::read_to_string(&dest).unwrap();
        assert!(content.contains("Name"));
        assert!(content.contains("Alice"));
    }

    #[test]
    fn test_export_csv_to_stdout() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "test.xlsx");
        setup_test_data(&file_path);

        let result = export_csv(&file_path, "-", None, ',', &default_global());
        assert!(result.is_ok());
    }

    #[test]
    fn test_export_json() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "test.xlsx");
        setup_test_data(&file_path);

        let dest = temp_dir.path().join("output.json");
        let result = export_json(
            &file_path,
            &dest.to_string_lossy(),
            None,
            false,
            &default_global(),
        );
        assert!(result.is_ok());
        assert!(dest.exists());
    }

    #[test]
    fn test_export_json_with_header() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "test.xlsx");
        setup_test_data(&file_path);

        let dest = temp_dir.path().join("output.json");
        let result = export_json(
            &file_path,
            &dest.to_string_lossy(),
            None,
            true,
            &default_global(),
        );
        assert!(result.is_ok());

        let content = std::fs::read_to_string(&dest).unwrap();
        // With header=true, keys should be "Name" and "Age"
        assert!(content.contains("\"Name\""));
    }

    #[test]
    fn test_export_tsv() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "test.xlsx");
        setup_test_data(&file_path);

        let dest = temp_dir.path().join("output.tsv");
        let result = export_tsv(&file_path, &dest.to_string_lossy(), None, &default_global());
        assert!(result.is_ok());
        assert!(dest.exists());

        let content = std::fs::read_to_string(&dest).unwrap();
        assert!(content.contains('\t'));
    }

    #[test]
    fn test_export_yaml() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "test.xlsx");
        setup_test_data(&file_path);

        let dest = temp_dir.path().join("output.yaml");
        let result = export_yaml(&file_path, &dest.to_string_lossy(), None, &default_global());
        assert!(result.is_ok());
        assert!(dest.exists());
    }

    #[test]
    fn test_export_markdown() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "test.xlsx");
        setup_test_data(&file_path);

        let dest = temp_dir.path().join("output.md");
        let result = export_markdown(&file_path, &dest.to_string_lossy(), None, &default_global());
        assert!(result.is_ok());
        assert!(dest.exists());

        let content = std::fs::read_to_string(&dest).unwrap();
        // Markdown table should have pipes and separators
        assert!(content.contains("|"));
        assert!(content.contains("---"));
    }

    #[test]
    fn test_export_ndjson() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "test.xlsx");
        setup_test_data(&file_path);

        let dest = temp_dir.path().join("output.ndjson");
        let result = export_ndjson(
            &file_path,
            &dest.to_string_lossy(),
            None,
            false,
            &default_global(),
        );
        assert!(result.is_ok());
        assert!(dest.exists());
    }

    #[test]
    fn test_export_ndjson_with_header() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "test.xlsx");
        setup_test_data(&file_path);

        let dest = temp_dir.path().join("output.ndjson");
        let result = export_ndjson(
            &file_path,
            &dest.to_string_lossy(),
            None,
            true,
            &default_global(),
        );
        assert!(result.is_ok());
    }

    #[test]
    fn test_export_meta() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "test.xlsx");
        setup_test_data(&file_path);

        let dest = temp_dir.path().join("meta.json");
        let result = export_meta(&file_path, &dest.to_string_lossy(), &default_global());
        assert!(result.is_ok());
        assert!(dest.exists());

        let content = std::fs::read_to_string(&dest).unwrap();
        assert!(content.contains("sheets"));
        assert!(content.contains("Sheet1"));
    }

    #[test]
    fn test_cell_to_json() {
        assert_eq!(cell_to_json(&CellValue::Empty), serde_json::Value::Null);
        assert_eq!(
            cell_to_json(&CellValue::String("test".to_string())),
            serde_json::Value::String("test".to_string())
        );
        assert_eq!(
            cell_to_json(&CellValue::Number(42.0)),
            serde_json::json!(42.0)
        );
        assert_eq!(
            cell_to_json(&CellValue::Boolean(true)),
            serde_json::Value::Bool(true)
        );
    }

    // Additional tests for better coverage

    #[test]
    fn test_export_csv_dry_run() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "test_dry.xlsx");
        setup_test_data(&file_path);

        let dest = temp_dir.path().join("output_dry.csv");
        let mut global = default_global();
        global.dry_run = true;

        let result = export_csv(&file_path, &dest.to_string_lossy(), None, ',', &global);
        assert!(result.is_ok());
        // Note: export functions don't check dry_run - they always export
        // dry_run is only checked in run() dispatcher for some commands
    }

    #[test]
    fn test_export_json_dry_run() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "test_json_dry.xlsx");
        setup_test_data(&file_path);

        let dest = temp_dir.path().join("output_dry.json");
        let mut global = default_global();
        global.dry_run = true;

        let result = export_json(&file_path, &dest.to_string_lossy(), None, false, &global);
        assert!(result.is_ok());
        // Note: export functions don't check dry_run - they always export
    }

    #[test]
    fn test_export_json_to_stdout() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "test_stdout.xlsx");
        setup_test_data(&file_path);

        let result = export_json(&file_path, "-", None, false, &default_global());
        assert!(result.is_ok());
    }

    #[test]
    fn test_export_yaml_to_stdout() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "test_yaml_stdout.xlsx");
        setup_test_data(&file_path);

        let result = export_yaml(&file_path, "-", None, &default_global());
        assert!(result.is_ok());
    }

    #[test]
    fn test_export_markdown_to_stdout() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "test_md_stdout.xlsx");
        setup_test_data(&file_path);

        let result = export_markdown(&file_path, "-", None, &default_global());
        assert!(result.is_ok());
    }

    #[test]
    fn test_export_ndjson_to_stdout() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "test_ndjson_stdout.xlsx");
        setup_test_data(&file_path);

        let result = export_ndjson(&file_path, "-", None, false, &default_global());
        assert!(result.is_ok());
    }

    #[test]
    fn test_export_meta_to_stdout() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "test_meta_stdout.xlsx");
        setup_test_data(&file_path);

        let result = export_meta(&file_path, "-", &default_global());
        assert!(result.is_ok());
    }

    #[test]
    fn test_export_all_csv() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "test_all.xlsx");
        setup_test_data(&file_path);

        let dest = temp_dir.path().join("output_all.csv");
        let result = export_all_csv(&file_path, &dest.to_string_lossy(), ',', &default_global());
        assert!(result.is_ok());
    }

    #[test]
    fn test_export_all_json() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "test_all_json.xlsx");
        setup_test_data(&file_path);

        let dest = temp_dir.path().join("output_all.json");
        let result = export_all_json(
            &file_path,
            &dest.to_string_lossy(),
            false,
            &default_global(),
        );
        assert!(result.is_ok());
    }

    #[test]
    fn test_export_all_yaml() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "test_all_yaml.xlsx");
        setup_test_data(&file_path);

        let dest = temp_dir.path().join("output_all.yaml");
        let result = export_all_yaml(&file_path, &dest.to_string_lossy(), &default_global());
        assert!(result.is_ok());
    }

    #[test]
    fn test_export_all_markdown() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "test_all_md.xlsx");
        setup_test_data(&file_path);

        let dest = temp_dir.path().join("output_all.md");
        let result = export_all_markdown(&file_path, &dest.to_string_lossy(), &default_global());
        assert!(result.is_ok());
    }

    #[test]
    fn test_export_all_ndjson() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "test_all_ndjson.xlsx");
        setup_test_data(&file_path);

        let dest = temp_dir.path().join("output_all.ndjson");
        let result = export_all_ndjson(
            &file_path,
            &dest.to_string_lossy(),
            false,
            &default_global(),
        );
        assert!(result.is_ok());
    }

    #[test]
    fn test_run_csv_command() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "run_csv.xlsx");
        setup_test_data(&file_path);

        let dest = temp_dir.path().join("run_output.csv");
        let args = ExportArgs {
            command: ExportCommand::Csv {
                source: file_path,
                dest: dest.to_string_lossy().to_string(),
                sheet: None,
                delimiter: ',',
                all: false,
            },
        };

        let result = run(&args, &default_global());
        assert!(result.is_ok());
    }

    #[test]
    fn test_run_json_command() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "run_json.xlsx");
        setup_test_data(&file_path);

        let dest = temp_dir.path().join("run_output.json");
        let args = ExportArgs {
            command: ExportCommand::Json {
                source: file_path,
                dest: dest.to_string_lossy().to_string(),
                sheet: None,
                header: false,
                all: false,
            },
        };

        let result = run(&args, &default_global());
        assert!(result.is_ok());
    }

    #[test]
    fn test_run_tsv_command() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "run_tsv.xlsx");
        setup_test_data(&file_path);

        let dest = temp_dir.path().join("run_output.tsv");
        let args = ExportArgs {
            command: ExportCommand::Tsv {
                source: file_path,
                dest: dest.to_string_lossy().to_string(),
                sheet: None,
                all: false,
            },
        };

        let result = run(&args, &default_global());
        assert!(result.is_ok());
    }

    #[test]
    fn test_run_yaml_command() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "run_yaml.xlsx");
        setup_test_data(&file_path);

        let dest = temp_dir.path().join("run_output.yaml");
        let args = ExportArgs {
            command: ExportCommand::Yaml {
                source: file_path,
                dest: dest.to_string_lossy().to_string(),
                sheet: None,
                all: false,
            },
        };

        let result = run(&args, &default_global());
        assert!(result.is_ok());
    }

    #[test]
    fn test_run_markdown_command() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "run_md.xlsx");
        setup_test_data(&file_path);

        let dest = temp_dir.path().join("run_output.md");
        let args = ExportArgs {
            command: ExportCommand::Markdown {
                source: file_path,
                dest: dest.to_string_lossy().to_string(),
                sheet: None,
                all: false,
            },
        };

        let result = run(&args, &default_global());
        assert!(result.is_ok());
    }

    #[test]
    fn test_run_ndjson_command() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "run_ndjson.xlsx");
        setup_test_data(&file_path);

        let dest = temp_dir.path().join("run_output.ndjson");
        let args = ExportArgs {
            command: ExportCommand::Ndjson {
                source: file_path,
                dest: dest.to_string_lossy().to_string(),
                sheet: None,
                header: false,
                all: false,
            },
        };

        let result = run(&args, &default_global());
        assert!(result.is_ok());
    }

    #[test]
    fn test_run_meta_command() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "run_meta.xlsx");
        setup_test_data(&file_path);

        let dest = temp_dir.path().join("run_meta.json");
        let args = ExportArgs {
            command: ExportCommand::Meta {
                source: file_path,
                dest: dest.to_string_lossy().to_string(),
            },
        };

        let result = run(&args, &default_global());
        assert!(result.is_ok());
    }

    #[test]
    fn test_run_csv_all_command() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "run_csv_all.xlsx");
        setup_test_data(&file_path);

        let dest = temp_dir.path().join("run_all.csv");
        let args = ExportArgs {
            command: ExportCommand::Csv {
                source: file_path,
                dest: dest.to_string_lossy().to_string(),
                sheet: None,
                delimiter: ',',
                all: true,
            },
        };

        let result = run(&args, &default_global());
        assert!(result.is_ok());
    }

    #[test]
    fn test_run_tsv_all_command() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "run_tsv_all.xlsx");
        setup_test_data(&file_path);

        let dest = temp_dir.path().join("run_all.tsv");
        let args = ExportArgs {
            command: ExportCommand::Tsv {
                source: file_path,
                dest: dest.to_string_lossy().to_string(),
                sheet: None,
                all: true,
            },
        };

        let result = run(&args, &default_global());
        assert!(result.is_ok());
    }

    #[test]
    fn test_cell_to_json_formula() {
        let value = CellValue::Formula {
            formula: "SUM(A1:A10)".to_string(),
            cached_result: Some(Box::new(CellValue::Number(100.0))),
        };
        let result = cell_to_json(&value);
        // cell_to_json uses to_display_string() for formulas, which returns "=formula"
        assert_eq!(
            result,
            serde_json::Value::String("=SUM(A1:A10)".to_string())
        );
    }

    #[test]
    fn test_cell_to_json_formula_no_cache() {
        let value = CellValue::Formula {
            formula: "SUM(A1:A10)".to_string(),
            cached_result: None,
        };
        let result = cell_to_json(&value);
        // cell_to_json uses to_display_string() for formulas
        assert_eq!(
            result,
            serde_json::Value::String("=SUM(A1:A10)".to_string())
        );
    }

    #[test]
    fn test_export_csv_with_sheet() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "sheet_csv.xlsx");
        setup_test_data(&file_path);

        let dest = temp_dir.path().join("sheet_output.csv");
        let result = export_csv(
            &file_path,
            &dest.to_string_lossy(),
            Some("Sheet1"),
            ',',
            &default_global(),
        );
        assert!(result.is_ok());
    }

    #[test]
    fn test_export_csv_semicolon_delimiter() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "semi_csv.xlsx");
        setup_test_data(&file_path);

        let dest = temp_dir.path().join("semi_output.csv");
        let result = export_csv(
            &file_path,
            &dest.to_string_lossy(),
            None,
            ';',
            &default_global(),
        );
        assert!(result.is_ok());

        let content = std::fs::read_to_string(&dest).unwrap();
        assert!(content.contains(";"));
    }

    #[test]
    fn test_export_tsv_to_stdout() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "tsv_stdout.xlsx");
        setup_test_data(&file_path);

        let result = export_tsv(&file_path, "-", None, &default_global());
        assert!(result.is_ok());
    }

    #[test]
    fn test_run_yaml_all_verbose() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "run_yaml_all.xlsx");
        setup_test_data(&file_path);

        let dest = temp_dir.path().join("run_yaml_all_output.yaml");
        let args = ExportArgs {
            command: ExportCommand::Yaml {
                source: file_path,
                dest: dest.to_string_lossy().to_string(),
                sheet: None,
                all: true,
            },
        };

        let result = run(&args, &default_global());
        assert!(result.is_ok());
    }

    #[test]
    fn test_run_markdown_all_verbose() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "run_md_all.xlsx");
        setup_test_data(&file_path);

        let dest = temp_dir.path().join("run_md_all_output.md");
        let args = ExportArgs {
            command: ExportCommand::Markdown {
                source: file_path,
                dest: dest.to_string_lossy().to_string(),
                sheet: None,
                all: true,
            },
        };

        let result = run(&args, &default_global());
        assert!(result.is_ok());
    }

    #[test]
    fn test_run_ndjson_all_verbose() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "run_ndjson_all.xlsx");
        setup_test_data(&file_path);

        let dest = temp_dir.path().join("run_ndjson_all_output.ndjson");
        let args = ExportArgs {
            command: ExportCommand::Ndjson {
                source: file_path,
                dest: dest.to_string_lossy().to_string(),
                sheet: None,
                header: false,
                all: true,
            },
        };

        let result = run(&args, &default_global());
        assert!(result.is_ok());
    }

    #[test]
    fn test_export_csv_sheet_not_found() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "sheet_nf.xlsx");

        let dest = temp_dir.path().join("sheet_nf_output.csv");
        let result = export_csv(
            &file_path,
            &dest.to_string_lossy(),
            Some("NonexistentSheet"),
            ',',
            &default_global(),
        );
        assert!(result.is_err());
    }

    #[test]
    fn test_export_json_sheet_not_found() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "json_sheet_nf.xlsx");

        let dest = temp_dir.path().join("json_sheet_nf.json");
        let result = export_json(
            &file_path,
            &dest.to_string_lossy(),
            Some("NonexistentSheet"),
            false,
            &default_global(),
        );
        assert!(result.is_err());
    }

    #[test]
    fn test_export_tsv_sheet_not_found() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "tsv_sheet_nf.xlsx");

        let dest = temp_dir.path().join("tsv_sheet_nf.tsv");
        let result = export_tsv(
            &file_path,
            &dest.to_string_lossy(),
            Some("NonexistentSheet"),
            &default_global(),
        );
        assert!(result.is_err());
    }

    #[test]
    fn test_export_yaml_sheet_not_found() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "yaml_sheet_nf.xlsx");

        let dest = temp_dir.path().join("yaml_sheet_nf.yaml");
        let result = export_yaml(
            &file_path,
            &dest.to_string_lossy(),
            Some("NonexistentSheet"),
            &default_global(),
        );
        assert!(result.is_err());
    }

    #[test]
    fn test_export_markdown_sheet_not_found() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "md_sheet_nf.xlsx");

        let dest = temp_dir.path().join("md_sheet_nf.md");
        let result = export_markdown(
            &file_path,
            &dest.to_string_lossy(),
            Some("NonexistentSheet"),
            &default_global(),
        );
        assert!(result.is_err());
    }

    #[test]
    fn test_export_ndjson_sheet_not_found() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "ndjson_sheet_nf.xlsx");

        let dest = temp_dir.path().join("ndjson_sheet_nf.ndjson");
        let result = export_ndjson(
            &file_path,
            &dest.to_string_lossy(),
            Some("NonexistentSheet"),
            false,
            &default_global(),
        );
        assert!(result.is_err());
    }

    #[test]
    fn test_cell_to_json_datetime() {
        // DateTime in xlex_core is stored as f64 (Excel serial date)
        let value = CellValue::DateTime(44945.4375); // 2024-01-15 10:30
        let result = cell_to_json(&value);
        // DateTime is serialized via to_display_string() which returns a string
        assert!(result.is_string());
    }

    #[test]
    fn test_cell_to_json_error() {
        use xlex_core::cell::CellError;
        let value = CellValue::Error(CellError::Value);
        let result = cell_to_json(&value);
        // Error is serialized as string
        assert!(result.is_string());
    }
}
