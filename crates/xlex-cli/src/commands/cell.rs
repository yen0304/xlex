//! Cell operations.

use anyhow::Result;
use clap::{Parser, Subcommand};
use colored::Colorize;

use xlex_core::{CellRef, CellValue, Workbook};

use super::{GlobalOptions, OutputFormat};

/// Arguments for cell operations.
#[derive(Parser)]
pub struct CellArgs {
    #[command(subcommand)]
    pub command: CellCommand,
}

#[derive(Subcommand)]
pub enum CellCommand {
    /// Get cell value
    Get {
        /// Path to the xlsx file
        file: std::path::PathBuf,
        /// Sheet name
        sheet: String,
        /// Cell reference (e.g., A1, B2)
        cell: String,
    },
    /// Set cell value
    Set {
        /// Path to the xlsx file
        file: std::path::PathBuf,
        /// Sheet name
        sheet: String,
        /// Cell reference (e.g., A1, B2)
        cell: String,
        /// Value to set
        value: String,
        /// Value type (string, number, boolean, formula)
        #[arg(long, short = 't', default_value = "auto")]
        value_type: ValueType,
    },
    /// Set cell formula
    Formula {
        /// Path to the xlsx file
        file: std::path::PathBuf,
        /// Sheet name
        sheet: String,
        /// Cell reference (e.g., A1, B2)
        cell: String,
        /// Formula (without leading =)
        formula: String,
    },
    /// Clear cell
    Clear {
        /// Path to the xlsx file
        file: std::path::PathBuf,
        /// Sheet name
        sheet: String,
        /// Cell reference (e.g., A1, B2)
        cell: String,
    },
    /// Get cell type
    Type {
        /// Path to the xlsx file
        file: std::path::PathBuf,
        /// Sheet name
        sheet: String,
        /// Cell reference (e.g., A1, B2)
        cell: String,
    },
    /// Batch cell operations from stdin
    Batch {
        /// Path to the xlsx file
        file: std::path::PathBuf,
    },
    /// Cell comment operations
    Comment(CommentArgs),
    /// Cell hyperlink operations
    Link(LinkArgs),
}

/// Arguments for comment operations.
#[derive(Parser)]
pub struct CommentArgs {
    #[command(subcommand)]
    pub command: CommentCommand,
}

#[derive(Subcommand)]
pub enum CommentCommand {
    /// Get cell comment
    Get {
        /// Path to the xlsx file
        file: std::path::PathBuf,
        /// Sheet name
        sheet: String,
        /// Cell reference
        cell: String,
    },
    /// Set cell comment
    Set {
        /// Path to the xlsx file
        file: std::path::PathBuf,
        /// Sheet name
        sheet: String,
        /// Cell reference
        cell: String,
        /// Comment text
        text: String,
        /// Author name
        #[arg(long)]
        author: Option<String>,
    },
    /// Remove cell comment
    Remove {
        /// Path to the xlsx file
        file: std::path::PathBuf,
        /// Sheet name
        sheet: String,
        /// Cell reference
        cell: String,
    },
    /// List all comments in a sheet
    List {
        /// Path to the xlsx file
        file: std::path::PathBuf,
        /// Sheet name
        sheet: String,
    },
}

/// Arguments for hyperlink operations.
#[derive(Parser)]
pub struct LinkArgs {
    #[command(subcommand)]
    pub command: LinkCommand,
}

#[derive(Subcommand)]
pub enum LinkCommand {
    /// Get cell hyperlink
    Get {
        /// Path to the xlsx file
        file: std::path::PathBuf,
        /// Sheet name
        sheet: String,
        /// Cell reference
        cell: String,
    },
    /// Set cell hyperlink
    Set {
        /// Path to the xlsx file
        file: std::path::PathBuf,
        /// Sheet name
        sheet: String,
        /// Cell reference
        cell: String,
        /// URL or internal reference
        url: String,
        /// Display text
        #[arg(long)]
        text: Option<String>,
    },
    /// Remove cell hyperlink
    Remove {
        /// Path to the xlsx file
        file: std::path::PathBuf,
        /// Sheet name
        sheet: String,
        /// Cell reference
        cell: String,
    },
}

#[derive(Debug, Clone, Copy, Default, PartialEq, Eq, clap::ValueEnum)]
pub enum ValueType {
    #[default]
    Auto,
    String,
    Number,
    Boolean,
    Formula,
}

/// Run cell operations.
pub fn run(args: &CellArgs, global: &GlobalOptions) -> Result<()> {
    match &args.command {
        CellCommand::Get { file, sheet, cell } => get(file, sheet, cell, global),
        CellCommand::Set {
            file,
            sheet,
            cell,
            value,
            value_type,
        } => set(file, sheet, cell, value, *value_type, global),
        CellCommand::Formula {
            file,
            sheet,
            cell,
            formula,
        } => set_formula(file, sheet, cell, formula, global),
        CellCommand::Clear { file, sheet, cell } => clear(file, sheet, cell, global),
        CellCommand::Type { file, sheet, cell } => get_type(file, sheet, cell, global),
        CellCommand::Batch { file } => batch(file, global),
        CellCommand::Comment(args) => run_comment(args, global),
        CellCommand::Link(args) => run_link(args, global),
    }
}

fn get(file: &std::path::Path, sheet: &str, cell: &str, global: &GlobalOptions) -> Result<()> {
    let workbook = Workbook::open(file)?;
    let cell_ref = CellRef::parse(cell)?;
    let value = workbook.get_cell(sheet, &cell_ref)?;

    if global.format == OutputFormat::Json {
        let json = serde_json::json!({
            "cell": cell,
            "type": value.type_name(),
            "value": match &value {
                CellValue::Empty => serde_json::Value::Null,
                CellValue::String(s) => serde_json::Value::String(s.clone()),
                CellValue::Number(n) => serde_json::json!(n),
                CellValue::Boolean(b) => serde_json::Value::Bool(*b),
                CellValue::Formula { formula, .. } => serde_json::json!({
                    "formula": formula,
                }),
                CellValue::Error(e) => serde_json::Value::String(e.to_string()),
                CellValue::DateTime(d) => serde_json::json!(d),
            },
        });
        println!("{}", serde_json::to_string_pretty(&json)?);
    } else {
        println!("{}", value.to_display_string());
    }

    Ok(())
}

fn set(
    file: &std::path::Path,
    sheet: &str,
    cell: &str,
    value: &str,
    value_type: ValueType,
    global: &GlobalOptions,
) -> Result<()> {
    if global.dry_run {
        println!("Would set {} in {} to '{}'", cell, sheet, value);
        return Ok(());
    }

    let mut workbook = Workbook::open(file)?;
    let cell_ref = CellRef::parse(cell)?;

    let cell_value = match value_type {
        ValueType::Auto => parse_auto_value(value),
        ValueType::String => CellValue::String(value.to_string()),
        ValueType::Number => {
            let n: f64 = value
                .parse()
                .map_err(|_| xlex_core::XlexError::InvalidCellValue {
                    message: format!("Cannot parse '{}' as number", value),
                })?;
            CellValue::Number(n)
        }
        ValueType::Boolean => {
            let b = value.eq_ignore_ascii_case("true")
                || value == "1"
                || value.eq_ignore_ascii_case("yes");
            CellValue::Boolean(b)
        }
        ValueType::Formula => CellValue::formula(value),
    };

    workbook.set_cell(sheet, cell_ref, cell_value)?;
    workbook.save()?;

    if !global.quiet {
        if global.format == OutputFormat::Json {
            let json = serde_json::json!({
                "action": "set",
                "cell": cell,
                "value": value,
            });
            println!("{}", serde_json::to_string_pretty(&json)?);
        } else {
            println!("Set {} to '{}'", cell.cyan(), value.green());
        }
    }

    Ok(())
}

fn set_formula(
    file: &std::path::Path,
    sheet: &str,
    cell: &str,
    formula: &str,
    global: &GlobalOptions,
) -> Result<()> {
    // Strip leading '=' if present (user may have included it)
    let formula = formula.strip_prefix('=').unwrap_or(formula);

    if global.dry_run {
        println!(
            "Would set formula in {} in {} to '={}'",
            cell, sheet, formula
        );
        return Ok(());
    }

    let mut workbook = Workbook::open(file)?;
    let cell_ref = CellRef::parse(cell)?;

    workbook.set_cell(sheet, cell_ref, CellValue::formula(formula))?;
    workbook.save()?;

    if !global.quiet {
        if global.format == OutputFormat::Json {
            let json = serde_json::json!({
                "action": "formula",
                "cell": cell,
                "formula": formula,
            });
            println!("{}", serde_json::to_string_pretty(&json)?);
        } else {
            println!("Set formula at {} to '={}'", cell.cyan(), formula.green());
        }
    }

    Ok(())
}

fn clear(file: &std::path::Path, sheet: &str, cell: &str, global: &GlobalOptions) -> Result<()> {
    if global.dry_run {
        println!("Would clear {} in {}", cell, sheet);
        return Ok(());
    }

    let mut workbook = Workbook::open(file)?;
    let cell_ref = CellRef::parse(cell)?;

    workbook.clear_cell(sheet, &cell_ref)?;
    workbook.save()?;

    if !global.quiet {
        if global.format == OutputFormat::Json {
            let json = serde_json::json!({
                "action": "clear",
                "cell": cell,
            });
            println!("{}", serde_json::to_string_pretty(&json)?);
        } else {
            println!("Cleared {}", cell.cyan());
        }
    }

    Ok(())
}

fn get_type(file: &std::path::Path, sheet: &str, cell: &str, global: &GlobalOptions) -> Result<()> {
    let workbook = Workbook::open(file)?;
    let cell_ref = CellRef::parse(cell)?;
    let value = workbook.get_cell(sheet, &cell_ref)?;

    if global.format == OutputFormat::Json {
        let json = serde_json::json!({
            "cell": cell,
            "type": value.type_name(),
        });
        println!("{}", serde_json::to_string_pretty(&json)?);
    } else {
        println!("{}", value.type_name());
    }

    Ok(())
}

fn batch(file: &std::path::Path, global: &GlobalOptions) -> Result<()> {
    use std::io::{self, BufRead};

    if global.dry_run {
        println!("Would process batch operations from stdin");
        return Ok(());
    }

    let mut workbook = Workbook::open(file)?;
    let stdin = io::stdin();
    let mut line_num = 0;
    let mut success_count = 0;
    let mut error_count = 0;

    for line in stdin.lock().lines() {
        line_num += 1;
        let line = line?;
        let line = line.trim();

        // Skip empty lines and comments
        if line.is_empty() || line.starts_with('#') {
            continue;
        }

        // Parse line: SHEET CELL VALUE or SHEET CELL (for clear)
        let parts: Vec<&str> = line.splitn(3, char::is_whitespace).collect();

        match parts.len() {
            2 => {
                // Clear operation: SHEET CELL
                let sheet = parts[0];
                let cell = parts[1];
                match CellRef::parse(cell) {
                    Ok(cell_ref) => {
                        if let Err(e) = workbook.clear_cell(sheet, &cell_ref) {
                            if !global.quiet {
                                eprintln!("Line {}: Error clearing {}: {}", line_num, cell, e);
                            }
                            error_count += 1;
                        } else {
                            success_count += 1;
                        }
                    }
                    Err(e) => {
                        if !global.quiet {
                            eprintln!("Line {}: Invalid cell '{}': {}", line_num, cell, e);
                        }
                        error_count += 1;
                    }
                }
            }
            3 => {
                // Set operation: SHEET CELL VALUE
                let sheet = parts[0];
                let cell = parts[1];
                let value = parts[2];
                match CellRef::parse(cell) {
                    Ok(cell_ref) => {
                        let cell_value = parse_auto_value(value);
                        if let Err(e) = workbook.set_cell(sheet, cell_ref, cell_value) {
                            if !global.quiet {
                                eprintln!("Line {}: Error setting {}: {}", line_num, cell, e);
                            }
                            error_count += 1;
                        } else {
                            success_count += 1;
                        }
                    }
                    Err(e) => {
                        if !global.quiet {
                            eprintln!("Line {}: Invalid cell '{}': {}", line_num, cell, e);
                        }
                        error_count += 1;
                    }
                }
            }
            _ => {
                if !global.quiet {
                    eprintln!("Line {}: Invalid format: '{}'", line_num, line);
                }
                error_count += 1;
            }
        }
    }

    workbook.save()?;

    if global.format == OutputFormat::Json {
        let json = serde_json::json!({
            "success": success_count,
            "errors": error_count,
            "total": success_count + error_count,
        });
        println!("{}", serde_json::to_string_pretty(&json)?);
    } else if !global.quiet {
        println!(
            "Batch complete: {} succeeded, {} errors",
            success_count.to_string().green(),
            error_count.to_string().red()
        );
    }

    Ok(())
}

fn run_comment(args: &CommentArgs, global: &GlobalOptions) -> Result<()> {
    match &args.command {
        CommentCommand::Get { file, sheet, cell } => comment_get(file, sheet, cell, global),
        CommentCommand::Set {
            file,
            sheet,
            cell,
            text,
            author,
        } => comment_set(file, sheet, cell, text, author.as_deref(), global),
        CommentCommand::Remove { file, sheet, cell } => comment_remove(file, sheet, cell, global),
        CommentCommand::List { file, sheet } => comment_list(file, sheet, global),
    }
}

fn comment_get(
    file: &std::path::Path,
    sheet: &str,
    cell: &str,
    global: &GlobalOptions,
) -> Result<()> {
    let workbook = Workbook::open(file)?;
    let cell_ref = CellRef::parse(cell)?;

    let sheet_obj =
        workbook
            .get_sheet(sheet)
            .ok_or_else(|| xlex_core::XlexError::SheetNotFound {
                name: sheet.to_string(),
            })?;

    let comment = sheet_obj
        .get_cell(&cell_ref)
        .and_then(|c| c.comment.clone());

    if global.format == OutputFormat::Json {
        let json = serde_json::json!({
            "cell": cell,
            "comment": comment,
        });
        println!("{}", serde_json::to_string_pretty(&json)?);
    } else if let Some(text) = comment {
        println!("{}", text);
    } else {
        println!("(no comment)");
    }
    Ok(())
}

fn comment_set(
    file: &std::path::Path,
    sheet: &str,
    cell: &str,
    text: &str,
    _author: Option<&str>,
    global: &GlobalOptions,
) -> Result<()> {
    if global.dry_run {
        println!("Would set comment on {} to '{}'", cell, text);
        return Ok(());
    }

    let mut workbook = Workbook::open(file)?;
    let cell_ref = CellRef::parse(cell)?;

    {
        let sheet_obj =
            workbook
                .get_sheet_mut(sheet)
                .ok_or_else(|| xlex_core::XlexError::SheetNotFound {
                    name: sheet.to_string(),
                })?;
        sheet_obj.set_cell_comment(&cell_ref, Some(text.to_string()));
    }

    workbook.save()?;

    if !global.quiet {
        println!("{} Set comment on {}", "✓".green(), cell.cyan());
    }
    Ok(())
}

fn comment_remove(
    file: &std::path::Path,
    sheet: &str,
    cell: &str,
    global: &GlobalOptions,
) -> Result<()> {
    if global.dry_run {
        println!("Would remove comment from {}", cell);
        return Ok(());
    }

    let mut workbook = Workbook::open(file)?;
    let cell_ref = CellRef::parse(cell)?;

    {
        let sheet_obj =
            workbook
                .get_sheet_mut(sheet)
                .ok_or_else(|| xlex_core::XlexError::SheetNotFound {
                    name: sheet.to_string(),
                })?;
        sheet_obj.set_cell_comment(&cell_ref, None);
    }

    workbook.save()?;

    if !global.quiet {
        println!("{} Removed comment from {}", "✓".green(), cell.cyan());
    }
    Ok(())
}

fn comment_list(file: &std::path::Path, sheet: &str, global: &GlobalOptions) -> Result<()> {
    let workbook = Workbook::open(file)?;
    let sheet_obj =
        workbook
            .get_sheet(sheet)
            .ok_or_else(|| xlex_core::XlexError::SheetNotFound {
                name: sheet.to_string(),
            })?;

    let comments: Vec<_> = sheet_obj
        .cells()
        .filter_map(|c| {
            c.comment.as_ref().map(|text| {
                serde_json::json!({
                    "cell": c.reference.to_a1(),
                    "comment": text,
                })
            })
        })
        .collect();

    if global.format == OutputFormat::Json {
        println!(
            "{}",
            serde_json::to_string_pretty(&serde_json::json!({ "comments": comments }))?
        );
    } else if comments.is_empty() {
        println!("No comments found");
    } else {
        println!("{} comment(s):", comments.len());
        for c in comments {
            println!(
                "  {}: {}",
                c["cell"].as_str().unwrap().cyan(),
                c["comment"].as_str().unwrap()
            );
        }
    }
    Ok(())
}

fn run_link(args: &LinkArgs, global: &GlobalOptions) -> Result<()> {
    match &args.command {
        LinkCommand::Get { file, sheet, cell } => link_get(file, sheet, cell, global),
        LinkCommand::Set {
            file,
            sheet,
            cell,
            url,
            text,
        } => link_set(file, sheet, cell, url, text.as_deref(), global),
        LinkCommand::Remove { file, sheet, cell } => link_remove(file, sheet, cell, global),
    }
}

fn link_get(file: &std::path::Path, sheet: &str, cell: &str, global: &GlobalOptions) -> Result<()> {
    let workbook = Workbook::open(file)?;
    let cell_ref = CellRef::parse(cell)?;

    let sheet_obj =
        workbook
            .get_sheet(sheet)
            .ok_or_else(|| xlex_core::XlexError::SheetNotFound {
                name: sheet.to_string(),
            })?;

    let hyperlink = sheet_obj
        .get_cell(&cell_ref)
        .and_then(|c| c.hyperlink.clone());

    if global.format == OutputFormat::Json {
        let json = serde_json::json!({
            "cell": cell,
            "link": hyperlink,
        });
        println!("{}", serde_json::to_string_pretty(&json)?);
    } else if let Some(url) = hyperlink {
        println!("{}", url);
    } else {
        println!("(no hyperlink)");
    }
    Ok(())
}

fn link_set(
    file: &std::path::Path,
    sheet: &str,
    cell: &str,
    url: &str,
    _text: Option<&str>,
    global: &GlobalOptions,
) -> Result<()> {
    if global.dry_run {
        println!("Would set hyperlink on {} to '{}'", cell, url);
        return Ok(());
    }

    let mut workbook = Workbook::open(file)?;
    let cell_ref = CellRef::parse(cell)?;

    {
        let sheet_obj =
            workbook
                .get_sheet_mut(sheet)
                .ok_or_else(|| xlex_core::XlexError::SheetNotFound {
                    name: sheet.to_string(),
                })?;
        sheet_obj.set_cell_hyperlink(&cell_ref, Some(url.to_string()));
    }

    workbook.save()?;

    if !global.quiet {
        println!(
            "{} Set hyperlink on {} to {}",
            "✓".green(),
            cell.cyan(),
            url
        );
    }
    Ok(())
}

fn link_remove(
    file: &std::path::Path,
    sheet: &str,
    cell: &str,
    global: &GlobalOptions,
) -> Result<()> {
    if global.dry_run {
        println!("Would remove hyperlink from {}", cell);
        return Ok(());
    }

    let mut workbook = Workbook::open(file)?;
    let cell_ref = CellRef::parse(cell)?;

    {
        let sheet_obj =
            workbook
                .get_sheet_mut(sheet)
                .ok_or_else(|| xlex_core::XlexError::SheetNotFound {
                    name: sheet.to_string(),
                })?;
        sheet_obj.set_cell_hyperlink(&cell_ref, None);
    }

    workbook.save()?;

    if !global.quiet {
        println!("{} Removed hyperlink from {}", "✓".green(), cell.cyan());
    }
    Ok(())
}

/// Parse a value string and infer its type.
pub fn parse_auto_value(value: &str) -> CellValue {
    // Check if it's a formula
    if value.starts_with('=') {
        return CellValue::formula(&value[1..]);
    }

    // Check if it's a boolean
    if value.eq_ignore_ascii_case("true") {
        return CellValue::Boolean(true);
    }
    if value.eq_ignore_ascii_case("false") {
        return CellValue::Boolean(false);
    }

    // Check if it's a number
    if let Ok(n) = value.parse::<f64>() {
        return CellValue::Number(n);
    }

    // Default to string
    CellValue::String(value.to_string())
}

#[cfg(test)]
mod tests {
    use super::*;
    use tempfile::TempDir;

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

    #[test]
    fn test_parse_auto_value_formula() {
        let value = parse_auto_value("=SUM(A1:A10)");
        match value {
            CellValue::Formula { formula, .. } => {
                assert_eq!(formula, "SUM(A1:A10)");
            }
            _ => panic!("Expected Formula"),
        }
    }

    #[test]
    fn test_parse_auto_value_boolean_true() {
        assert_eq!(parse_auto_value("true"), CellValue::Boolean(true));
        assert_eq!(parse_auto_value("TRUE"), CellValue::Boolean(true));
        assert_eq!(parse_auto_value("True"), CellValue::Boolean(true));
    }

    #[test]
    fn test_parse_auto_value_boolean_false() {
        assert_eq!(parse_auto_value("false"), CellValue::Boolean(false));
        assert_eq!(parse_auto_value("FALSE"), CellValue::Boolean(false));
        assert_eq!(parse_auto_value("False"), CellValue::Boolean(false));
    }

    #[test]
    fn test_parse_auto_value_number() {
        assert_eq!(parse_auto_value("42"), CellValue::Number(42.0));
        assert_eq!(parse_auto_value("3.14"), CellValue::Number(3.14));
        assert_eq!(parse_auto_value("-100"), CellValue::Number(-100.0));
    }

    #[test]
    fn test_parse_auto_value_string() {
        assert_eq!(
            parse_auto_value("Hello"),
            CellValue::String("Hello".to_string())
        );
        assert_eq!(
            parse_auto_value("not a number"),
            CellValue::String("not a number".to_string())
        );
    }

    #[test]
    fn test_get_cell() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "get.xlsx");

        // Set a value first
        {
            let mut wb = Workbook::open(&file_path).unwrap();
            let cell_ref = CellRef::parse("A1").unwrap();
            wb.set_cell("Sheet1", cell_ref, CellValue::String("Test".to_string()))
                .unwrap();
            wb.save().unwrap();
        }

        let result = get(&file_path, "Sheet1", "A1", &default_global());
        assert!(result.is_ok());
    }

    #[test]
    fn test_get_cell_json() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "get_json.xlsx");

        let mut global = default_global();
        global.format = OutputFormat::Json;

        let result = get(&file_path, "Sheet1", "A1", &global);
        assert!(result.is_ok());
    }

    #[test]
    fn test_set_cell() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "set.xlsx");

        let result = set(
            &file_path,
            "Sheet1",
            "A1",
            "Hello",
            ValueType::String,
            &default_global(),
        );
        assert!(result.is_ok());

        let wb = Workbook::open(&file_path).unwrap();
        let cell_ref = CellRef::parse("A1").unwrap();
        let value = wb.get_cell("Sheet1", &cell_ref).unwrap();
        assert_eq!(value, CellValue::String("Hello".to_string()));
    }

    #[test]
    fn test_set_cell_number() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "set_num.xlsx");

        let result = set(
            &file_path,
            "Sheet1",
            "B2",
            "42",
            ValueType::Number,
            &default_global(),
        );
        assert!(result.is_ok());

        let wb = Workbook::open(&file_path).unwrap();
        let cell_ref = CellRef::parse("B2").unwrap();
        let value = wb.get_cell("Sheet1", &cell_ref).unwrap();
        assert_eq!(value, CellValue::Number(42.0));
    }

    #[test]
    fn test_set_cell_dry_run() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "set_dry.xlsx");

        let mut global = default_global();
        global.dry_run = true;

        let result = set(
            &file_path,
            "Sheet1",
            "A1",
            "Value",
            ValueType::Auto,
            &global,
        );
        assert!(result.is_ok());

        // Cell should be empty
        let wb = Workbook::open(&file_path).unwrap();
        let cell_ref = CellRef::parse("A1").unwrap();
        let value = wb.get_cell("Sheet1", &cell_ref).unwrap();
        assert_eq!(value, CellValue::Empty);
    }

    #[test]
    fn test_set_formula() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "formula.xlsx");

        let result = set_formula(&file_path, "Sheet1", "A1", "SUM(B1:B10)", &default_global());
        assert!(result.is_ok());

        let wb = Workbook::open(&file_path).unwrap();
        let cell_ref = CellRef::parse("A1").unwrap();
        let value = wb.get_cell("Sheet1", &cell_ref).unwrap();
        match value {
            CellValue::Formula { formula, .. } => {
                assert_eq!(formula, "SUM(B1:B10)");
            }
            _ => panic!("Expected formula"),
        }
    }

    #[test]
    fn test_set_formula_with_equals() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "formula_eq.xlsx");

        let result = set_formula(&file_path, "Sheet1", "A1", "=A1+B1", &default_global());
        assert!(result.is_ok());

        let wb = Workbook::open(&file_path).unwrap();
        let cell_ref = CellRef::parse("A1").unwrap();
        let value = wb.get_cell("Sheet1", &cell_ref).unwrap();
        match value {
            CellValue::Formula { formula, .. } => {
                assert_eq!(formula, "A1+B1"); // Equals should be stripped
            }
            _ => panic!("Expected formula"),
        }
    }

    #[test]
    fn test_clear_cell() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "clear.xlsx");

        // Set a value first
        set(
            &file_path,
            "Sheet1",
            "A1",
            "ToDelete",
            ValueType::String,
            &default_global(),
        )
        .unwrap();

        let result = clear(&file_path, "Sheet1", "A1", &default_global());
        assert!(result.is_ok());

        let wb = Workbook::open(&file_path).unwrap();
        let cell_ref = CellRef::parse("A1").unwrap();
        let value = wb.get_cell("Sheet1", &cell_ref).unwrap();
        assert_eq!(value, CellValue::Empty);
    }

    #[test]
    fn test_get_type() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "type.xlsx");

        set(
            &file_path,
            "Sheet1",
            "A1",
            "42",
            ValueType::Number,
            &default_global(),
        )
        .unwrap();

        let result = get_type(&file_path, "Sheet1", "A1", &default_global());
        assert!(result.is_ok());
    }

    #[test]
    fn test_comment_get() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "comment_get.xlsx");

        let result = comment_get(&file_path, "Sheet1", "A1", &default_global());
        assert!(result.is_ok());
    }

    #[test]
    fn test_comment_set() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "comment_set.xlsx");

        let result = comment_set(
            &file_path,
            "Sheet1",
            "A1",
            "This is a comment",
            None,
            &default_global(),
        );
        assert!(result.is_ok());
    }

    #[test]
    fn test_comment_remove() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "comment_rm.xlsx");

        // Set a comment first
        comment_set(
            &file_path,
            "Sheet1",
            "A1",
            "Comment",
            None,
            &default_global(),
        )
        .unwrap();

        let result = comment_remove(&file_path, "Sheet1", "A1", &default_global());
        assert!(result.is_ok());
    }

    #[test]
    fn test_comment_list() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "comment_list.xlsx");

        let result = comment_list(&file_path, "Sheet1", &default_global());
        assert!(result.is_ok());
    }

    #[test]
    fn test_link_get() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "link_get.xlsx");

        let result = link_get(&file_path, "Sheet1", "A1", &default_global());
        assert!(result.is_ok());
    }

    #[test]
    fn test_link_set() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "link_set.xlsx");

        let result = link_set(
            &file_path,
            "Sheet1",
            "A1",
            "https://example.com",
            None,
            &default_global(),
        );
        assert!(result.is_ok());
    }

    #[test]
    fn test_link_remove() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "link_rm.xlsx");

        // Set a link first
        link_set(
            &file_path,
            "Sheet1",
            "A1",
            "https://example.com",
            None,
            &default_global(),
        )
        .unwrap();

        let result = link_remove(&file_path, "Sheet1", "A1", &default_global());
        assert!(result.is_ok());
    }

    // Additional tests for better coverage

    #[test]
    fn test_set_cell_boolean() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "set_bool.xlsx");

        let result = set(
            &file_path,
            "Sheet1",
            "A1",
            "true",
            ValueType::Boolean,
            &default_global(),
        );
        assert!(result.is_ok());

        let wb = Workbook::open(&file_path).unwrap();
        let cell_ref = CellRef::parse("A1").unwrap();
        let value = wb.get_cell("Sheet1", &cell_ref).unwrap();
        assert_eq!(value, CellValue::Boolean(true));
    }

    #[test]
    fn test_set_cell_boolean_yes() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "set_bool_yes.xlsx");

        let result = set(
            &file_path,
            "Sheet1",
            "A1",
            "yes",
            ValueType::Boolean,
            &default_global(),
        );
        assert!(result.is_ok());

        let wb = Workbook::open(&file_path).unwrap();
        let cell_ref = CellRef::parse("A1").unwrap();
        let value = wb.get_cell("Sheet1", &cell_ref).unwrap();
        assert_eq!(value, CellValue::Boolean(true));
    }

    #[test]
    fn test_set_cell_formula_type() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "set_formula_type.xlsx");

        let result = set(
            &file_path,
            "Sheet1",
            "A1",
            "SUM(A2:A10)",
            ValueType::Formula,
            &default_global(),
        );
        assert!(result.is_ok());

        let wb = Workbook::open(&file_path).unwrap();
        let cell_ref = CellRef::parse("A1").unwrap();
        let value = wb.get_cell("Sheet1", &cell_ref).unwrap();
        match value {
            CellValue::Formula { formula, .. } => {
                assert_eq!(formula, "SUM(A2:A10)");
            }
            _ => panic!("Expected formula"),
        }
    }

    #[test]
    fn test_set_cell_auto_number() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "set_auto_num.xlsx");

        let result = set(
            &file_path,
            "Sheet1",
            "A1",
            "123.45",
            ValueType::Auto,
            &default_global(),
        );
        assert!(result.is_ok());

        let wb = Workbook::open(&file_path).unwrap();
        let cell_ref = CellRef::parse("A1").unwrap();
        let value = wb.get_cell("Sheet1", &cell_ref).unwrap();
        assert_eq!(value, CellValue::Number(123.45));
    }

    #[test]
    fn test_set_number_invalid() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "set_num_inv.xlsx");

        let result = set(
            &file_path,
            "Sheet1",
            "A1",
            "not_a_number",
            ValueType::Number,
            &default_global(),
        );
        assert!(result.is_err());
    }

    #[test]
    fn test_set_cell_json_output() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "set_json.xlsx");

        let mut global = default_global();
        global.format = OutputFormat::Json;
        global.quiet = false;

        let result = set(&file_path, "Sheet1", "A1", "Test", ValueType::Auto, &global);
        assert!(result.is_ok());
    }

    #[test]
    fn test_set_formula_dry_run() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "formula_dry.xlsx");

        let mut global = default_global();
        global.dry_run = true;

        let result = set_formula(&file_path, "Sheet1", "A1", "SUM(A1:A10)", &global);
        assert!(result.is_ok());
    }

    #[test]
    fn test_set_formula_json_output() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "formula_json.xlsx");

        let mut global = default_global();
        global.format = OutputFormat::Json;
        global.quiet = false;

        let result = set_formula(&file_path, "Sheet1", "A1", "A1+B1", &global);
        assert!(result.is_ok());
    }

    #[test]
    fn test_clear_dry_run() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "clear_dry.xlsx");

        let mut global = default_global();
        global.dry_run = true;

        let result = clear(&file_path, "Sheet1", "A1", &global);
        assert!(result.is_ok());
    }

    #[test]
    fn test_clear_json_output() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "clear_json.xlsx");

        let mut global = default_global();
        global.format = OutputFormat::Json;
        global.quiet = false;

        let result = clear(&file_path, "Sheet1", "A1", &global);
        assert!(result.is_ok());
    }

    #[test]
    fn test_get_type_json() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "type_json.xlsx");

        set(
            &file_path,
            "Sheet1",
            "A1",
            "42",
            ValueType::Number,
            &default_global(),
        )
        .unwrap();

        let mut global = default_global();
        global.format = OutputFormat::Json;

        let result = get_type(&file_path, "Sheet1", "A1", &global);
        assert!(result.is_ok());
    }

    #[test]
    fn test_comment_set_dry_run() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "comment_dry.xlsx");

        let mut global = default_global();
        global.dry_run = true;

        let result = comment_set(&file_path, "Sheet1", "A1", "Test", None, &global);
        assert!(result.is_ok());
    }

    #[test]
    fn test_comment_remove_dry_run() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "comment_rm_dry.xlsx");

        let mut global = default_global();
        global.dry_run = true;

        let result = comment_remove(&file_path, "Sheet1", "A1", &global);
        assert!(result.is_ok());
    }

    #[test]
    fn test_comment_list_json() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "comment_list_json.xlsx");

        let mut global = default_global();
        global.format = OutputFormat::Json;

        let result = comment_list(&file_path, "Sheet1", &global);
        assert!(result.is_ok());
    }

    #[test]
    fn test_comment_get_json() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "comment_get_json.xlsx");

        let mut global = default_global();
        global.format = OutputFormat::Json;

        let result = comment_get(&file_path, "Sheet1", "A1", &global);
        assert!(result.is_ok());
    }

    #[test]
    fn test_link_set_dry_run() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "link_dry.xlsx");

        let mut global = default_global();
        global.dry_run = true;

        let result = link_set(
            &file_path,
            "Sheet1",
            "A1",
            "https://example.com",
            None,
            &global,
        );
        assert!(result.is_ok());
    }

    #[test]
    fn test_link_remove_dry_run() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "link_rm_dry.xlsx");

        let mut global = default_global();
        global.dry_run = true;

        let result = link_remove(&file_path, "Sheet1", "A1", &global);
        assert!(result.is_ok());
    }

    #[test]
    fn test_link_get_json() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "link_get_json.xlsx");

        let mut global = default_global();
        global.format = OutputFormat::Json;

        let result = link_get(&file_path, "Sheet1", "A1", &global);
        assert!(result.is_ok());
    }

    #[test]
    fn test_run_get_command() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "run_get.xlsx");

        let args = CellArgs {
            command: CellCommand::Get {
                file: file_path,
                sheet: "Sheet1".to_string(),
                cell: "A1".to_string(),
            },
        };

        let result = run(&args, &default_global());
        assert!(result.is_ok());
    }

    #[test]
    fn test_run_set_command() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "run_set.xlsx");

        let args = CellArgs {
            command: CellCommand::Set {
                file: file_path,
                sheet: "Sheet1".to_string(),
                cell: "A1".to_string(),
                value: "Test".to_string(),
                value_type: ValueType::Auto,
            },
        };

        let result = run(&args, &default_global());
        assert!(result.is_ok());
    }

    #[test]
    fn test_run_formula_command() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "run_formula.xlsx");

        let args = CellArgs {
            command: CellCommand::Formula {
                file: file_path,
                sheet: "Sheet1".to_string(),
                cell: "A1".to_string(),
                formula: "SUM(B1:B10)".to_string(),
            },
        };

        let result = run(&args, &default_global());
        assert!(result.is_ok());
    }

    #[test]
    fn test_run_clear_command() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "run_clear.xlsx");

        let args = CellArgs {
            command: CellCommand::Clear {
                file: file_path,
                sheet: "Sheet1".to_string(),
                cell: "A1".to_string(),
            },
        };

        let result = run(&args, &default_global());
        assert!(result.is_ok());
    }

    #[test]
    fn test_run_type_command() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "run_type.xlsx");

        let args = CellArgs {
            command: CellCommand::Type {
                file: file_path,
                sheet: "Sheet1".to_string(),
                cell: "A1".to_string(),
            },
        };

        let result = run(&args, &default_global());
        assert!(result.is_ok());
    }

    #[test]
    fn test_parse_auto_value_empty() {
        assert_eq!(parse_auto_value(""), CellValue::String("".to_string()));
    }

    #[test]
    fn test_parse_auto_value_negative_decimal() {
        assert_eq!(parse_auto_value("-3.14"), CellValue::Number(-3.14));
    }

    #[test]
    fn test_get_cell_with_number() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "get_num.xlsx");

        {
            let mut wb = Workbook::open(&file_path).unwrap();
            let cell_ref = CellRef::parse("A1").unwrap();
            wb.set_cell("Sheet1", cell_ref, CellValue::Number(42.5))
                .unwrap();
            wb.save().unwrap();
        }

        let result = get(&file_path, "Sheet1", "A1", &default_global());
        assert!(result.is_ok());
    }

    #[test]
    fn test_get_cell_with_boolean() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "get_bool.xlsx");

        {
            let mut wb = Workbook::open(&file_path).unwrap();
            let cell_ref = CellRef::parse("A1").unwrap();
            wb.set_cell("Sheet1", cell_ref, CellValue::Boolean(true))
                .unwrap();
            wb.save().unwrap();
        }

        let result = get(&file_path, "Sheet1", "A1", &default_global());
        assert!(result.is_ok());
    }

    #[test]
    fn test_get_cell_with_formula() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "get_formula.xlsx");

        {
            let mut wb = Workbook::open(&file_path).unwrap();
            let cell_ref = CellRef::parse("A1").unwrap();
            wb.set_cell("Sheet1", cell_ref, CellValue::formula("SUM(B1:B10)"))
                .unwrap();
            wb.save().unwrap();
        }

        let mut global = default_global();
        global.format = OutputFormat::Json;

        let result = get(&file_path, "Sheet1", "A1", &global);
        assert!(result.is_ok());
    }

    #[test]
    fn test_get_cell_empty() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "get_empty.xlsx");

        let mut global = default_global();
        global.format = OutputFormat::Json;

        let result = get(&file_path, "Sheet1", "A1", &global);
        assert!(result.is_ok());
    }

    // run command tests
    #[test]
    fn test_run_comment_get_command() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "run_comment_get.xlsx");

        let args = CellArgs {
            command: CellCommand::Comment(CommentArgs {
                command: CommentCommand::Get {
                    file: file_path,
                    sheet: "Sheet1".to_string(),
                    cell: "A1".to_string(),
                },
            }),
        };

        let result = run(&args, &default_global());
        assert!(result.is_ok());
    }

    #[test]
    fn test_run_comment_set_command() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "run_comment_set.xlsx");

        let args = CellArgs {
            command: CellCommand::Comment(CommentArgs {
                command: CommentCommand::Set {
                    file: file_path,
                    sheet: "Sheet1".to_string(),
                    cell: "A1".to_string(),
                    text: "Test comment".to_string(),
                    author: Some("Author".to_string()),
                },
            }),
        };

        let result = run(&args, &default_global());
        assert!(result.is_ok());
    }

    #[test]
    fn test_run_comment_remove_command() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "run_comment_rm.xlsx");

        let args = CellArgs {
            command: CellCommand::Comment(CommentArgs {
                command: CommentCommand::Remove {
                    file: file_path,
                    sheet: "Sheet1".to_string(),
                    cell: "A1".to_string(),
                },
            }),
        };

        let result = run(&args, &default_global());
        assert!(result.is_ok());
    }

    #[test]
    fn test_run_comment_list_command() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "run_comment_list.xlsx");

        let args = CellArgs {
            command: CellCommand::Comment(CommentArgs {
                command: CommentCommand::List {
                    file: file_path,
                    sheet: "Sheet1".to_string(),
                },
            }),
        };

        let result = run(&args, &default_global());
        assert!(result.is_ok());
    }

    #[test]
    fn test_run_link_get_command() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "run_link_get.xlsx");

        let args = CellArgs {
            command: CellCommand::Link(LinkArgs {
                command: LinkCommand::Get {
                    file: file_path,
                    sheet: "Sheet1".to_string(),
                    cell: "A1".to_string(),
                },
            }),
        };

        let result = run(&args, &default_global());
        assert!(result.is_ok());
    }

    #[test]
    fn test_run_link_set_command() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "run_link_set.xlsx");

        let args = CellArgs {
            command: CellCommand::Link(LinkArgs {
                command: LinkCommand::Set {
                    file: file_path,
                    sheet: "Sheet1".to_string(),
                    cell: "A1".to_string(),
                    url: "https://example.com".to_string(),
                    text: Some("Example".to_string()),
                },
            }),
        };

        let result = run(&args, &default_global());
        assert!(result.is_ok());
    }

    #[test]
    fn test_run_link_remove_command() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "run_link_rm.xlsx");

        let args = CellArgs {
            command: CellCommand::Link(LinkArgs {
                command: LinkCommand::Remove {
                    file: file_path,
                    sheet: "Sheet1".to_string(),
                    cell: "A1".to_string(),
                },
            }),
        };

        let result = run(&args, &default_global());
        assert!(result.is_ok());
    }

    // JSON output tests
    #[test]
    fn test_get_type_json_format() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "type_json.xlsx");

        let mut global = default_global();
        global.format = OutputFormat::Json;

        let result = get_type(&file_path, "Sheet1", "A1", &global);
        assert!(result.is_ok());
    }

    #[test]
    fn test_comment_get_json_format() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "comment_json2.xlsx");

        let mut global = default_global();
        global.format = OutputFormat::Json;

        let result = comment_get(&file_path, "Sheet1", "A1", &global);
        assert!(result.is_ok());
    }

    #[test]
    fn test_comment_list_json_format() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "comment_list_json2.xlsx");

        let mut global = default_global();
        global.format = OutputFormat::Json;

        let result = comment_list(&file_path, "Sheet1", &global);
        assert!(result.is_ok());
    }

    #[test]
    fn test_link_get_json_format() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "link_json2.xlsx");

        let mut global = default_global();
        global.format = OutputFormat::Json;

        let result = link_get(&file_path, "Sheet1", "A1", &global);
        assert!(result.is_ok());
    }

    // Verbose output tests
    #[test]
    fn test_set_verbose_json() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "set_verbose.xlsx");

        let mut global = default_global();
        global.quiet = false;
        global.format = OutputFormat::Json;

        let result = set(
            &file_path,
            "Sheet1",
            "A1",
            "Value",
            ValueType::String,
            &global,
        );
        assert!(result.is_ok());
    }

    #[test]
    fn test_set_verbose_text() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "set_verbose_txt.xlsx");

        let mut global = default_global();
        global.quiet = false;
        global.format = OutputFormat::Text;

        let result = set(
            &file_path,
            "Sheet1",
            "A1",
            "Value",
            ValueType::String,
            &global,
        );
        assert!(result.is_ok());
    }

    #[test]
    fn test_set_formula_verbose_json() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "formula_verbose.xlsx");

        let mut global = default_global();
        global.quiet = false;
        global.format = OutputFormat::Json;

        let result = set_formula(&file_path, "Sheet1", "A1", "SUM(A1:A10)", &global);
        assert!(result.is_ok());
    }

    #[test]
    fn test_set_formula_verbose_text() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "formula_verbose_txt.xlsx");

        let mut global = default_global();
        global.quiet = false;
        global.format = OutputFormat::Text;

        let result = set_formula(&file_path, "Sheet1", "A1", "SUM(A1:A10)", &global);
        assert!(result.is_ok());
    }

    #[test]
    fn test_clear_verbose_json() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "clear_verbose.xlsx");

        let mut global = default_global();
        global.quiet = false;
        global.format = OutputFormat::Json;

        let result = clear(&file_path, "Sheet1", "A1", &global);
        assert!(result.is_ok());
    }

    #[test]
    fn test_clear_verbose_text() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "clear_verbose_txt.xlsx");

        let mut global = default_global();
        global.quiet = false;
        global.format = OutputFormat::Text;

        let result = clear(&file_path, "Sheet1", "A1", &global);
        assert!(result.is_ok());
    }

    #[test]
    fn test_comment_set_verbose() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "comment_verbose.xlsx");

        let mut global = default_global();
        global.quiet = false;

        let result = comment_set(&file_path, "Sheet1", "A1", "Comment", None, &global);
        assert!(result.is_ok());
    }

    #[test]
    fn test_comment_remove_verbose() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "comment_rm_verbose.xlsx");

        let mut global = default_global();
        global.quiet = false;

        // First set a comment
        comment_set(
            &file_path,
            "Sheet1",
            "A1",
            "Comment",
            None,
            &default_global(),
        )
        .unwrap();

        let result = comment_remove(&file_path, "Sheet1", "A1", &global);
        assert!(result.is_ok());
    }

    #[test]
    fn test_comment_list_with_comments() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "comment_list_with.xlsx");

        // Set a comment first
        comment_set(
            &file_path,
            "Sheet1",
            "A1",
            "Comment",
            None,
            &default_global(),
        )
        .unwrap();

        let mut global = default_global();
        global.quiet = false;

        let result = comment_list(&file_path, "Sheet1", &global);
        assert!(result.is_ok());
    }

    #[test]
    fn test_link_set_verbose() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "link_verbose.xlsx");

        let mut global = default_global();
        global.quiet = false;

        let result = link_set(
            &file_path,
            "Sheet1",
            "A1",
            "https://example.com",
            None,
            &global,
        );
        assert!(result.is_ok());
    }

    #[test]
    fn test_link_remove_verbose() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "link_rm_verbose.xlsx");

        // First set a link
        link_set(
            &file_path,
            "Sheet1",
            "A1",
            "https://example.com",
            None,
            &default_global(),
        )
        .unwrap();

        let mut global = default_global();
        global.quiet = false;

        let result = link_remove(&file_path, "Sheet1", "A1", &global);
        assert!(result.is_ok());
    }

    #[test]
    fn test_link_get_with_link() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "link_get_with.xlsx");

        // First set a link
        link_set(
            &file_path,
            "Sheet1",
            "A1",
            "https://example.com",
            None,
            &default_global(),
        )
        .unwrap();

        let mut global = default_global();
        global.quiet = false;

        let result = link_get(&file_path, "Sheet1", "A1", &global);
        assert!(result.is_ok());
    }

    #[test]
    fn test_comment_get_with_comment() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "comment_get_with.xlsx");

        // First set a comment
        comment_set(
            &file_path,
            "Sheet1",
            "A1",
            "Test Comment",
            None,
            &default_global(),
        )
        .unwrap();

        let mut global = default_global();
        global.quiet = false;

        let result = comment_get(&file_path, "Sheet1", "A1", &global);
        assert!(result.is_ok());
    }

    #[test]
    fn test_batch_dry_run() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "batch_dry.xlsx");

        let mut global = default_global();
        global.dry_run = true;

        let result = batch(&file_path, &global);
        assert!(result.is_ok());
    }

    #[test]
    fn test_get_display_text_string() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "get_display.xlsx");

        {
            let mut wb = Workbook::open(&file_path).unwrap();
            let cell_ref = CellRef::parse("A1").unwrap();
            wb.set_cell(
                "Sheet1",
                cell_ref,
                CellValue::String("Hello World".to_string()),
            )
            .unwrap();
            wb.save().unwrap();
        }

        let mut global = default_global();
        global.format = OutputFormat::Text;

        let result = get(&file_path, "Sheet1", "A1", &global);
        assert!(result.is_ok());
    }

    #[test]
    fn test_set_cell_error_invalid_number() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "set_invalid_num.xlsx");

        let result = set(
            &file_path,
            "Sheet1",
            "A1",
            "not_a_number",
            ValueType::Number,
            &default_global(),
        );
        assert!(result.is_err());
    }

    #[test]
    fn test_set_cell_verbose_output() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "set_verbose.xlsx");

        let mut global = default_global();
        global.quiet = false;

        let result = set(
            &file_path,
            "Sheet1",
            "A1",
            "Test",
            ValueType::String,
            &global,
        );
        assert!(result.is_ok());
    }

    #[test]
    fn test_set_formula_verbose() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "formula_verbose.xlsx");

        let mut global = default_global();
        global.quiet = false;

        let result = set_formula(&file_path, "Sheet1", "A1", "A2+B2", &global);
        assert!(result.is_ok());
    }

    #[test]
    fn test_clear_cell_verbose() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "clear_verbose.xlsx");

        // First set a value
        set(
            &file_path,
            "Sheet1",
            "A1",
            "Value",
            ValueType::String,
            &default_global(),
        )
        .unwrap();

        let mut global = default_global();
        global.quiet = false;

        let result = clear(&file_path, "Sheet1", "A1", &global);
        assert!(result.is_ok());
    }

    #[test]
    fn test_get_type_verbose() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "type_verbose.xlsx");

        set(
            &file_path,
            "Sheet1",
            "A1",
            "Hello",
            ValueType::String,
            &default_global(),
        )
        .unwrap();

        let mut global = default_global();
        global.quiet = false;

        let result = get_type(&file_path, "Sheet1", "A1", &global);
        assert!(result.is_ok());
    }
}
