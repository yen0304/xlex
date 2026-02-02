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

fn get(
    file: &std::path::Path,
    sheet: &str,
    cell: &str,
    global: &GlobalOptions,
) -> Result<()> {
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
            let n: f64 = value.parse().map_err(|_| {
                xlex_core::XlexError::InvalidCellValue {
                    message: format!("Cannot parse '{}' as number", value),
                }
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
        println!("Would set formula in {} in {} to '={}'", cell, sheet, formula);
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

fn clear(
    file: &std::path::Path,
    sheet: &str,
    cell: &str,
    global: &GlobalOptions,
) -> Result<()> {
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

fn get_type(
    file: &std::path::Path,
    sheet: &str,
    cell: &str,
    global: &GlobalOptions,
) -> Result<()> {
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
    
    let sheet_obj = workbook
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
        let sheet_obj = workbook
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
        let sheet_obj = workbook
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

fn comment_list(
    file: &std::path::Path,
    sheet: &str,
    global: &GlobalOptions,
) -> Result<()> {
    let workbook = Workbook::open(file)?;
    let sheet_obj = workbook
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
        println!("{}", serde_json::to_string_pretty(&serde_json::json!({ "comments": comments }))?);
    } else if comments.is_empty() {
        println!("No comments found");
    } else {
        println!("{} comment(s):", comments.len());
        for c in comments {
            println!("  {}: {}", c["cell"].as_str().unwrap().cyan(), c["comment"].as_str().unwrap());
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

fn link_get(
    file: &std::path::Path,
    sheet: &str,
    cell: &str,
    global: &GlobalOptions,
) -> Result<()> {
    let workbook = Workbook::open(file)?;
    let cell_ref = CellRef::parse(cell)?;

    let sheet_obj = workbook
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
        let sheet_obj = workbook
            .get_sheet_mut(sheet)
            .ok_or_else(|| xlex_core::XlexError::SheetNotFound {
                name: sheet.to_string(),
            })?;
        sheet_obj.set_cell_hyperlink(&cell_ref, Some(url.to_string()));
    }

    workbook.save()?;

    if !global.quiet {
        println!("{} Set hyperlink on {} to {}", "✓".green(), cell.cyan(), url);
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
        let sheet_obj = workbook
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
