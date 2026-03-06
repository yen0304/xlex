//! Global search operations - search across all sheets in a workbook.

use anyhow::Result;
use clap::Parser;
use colored::Colorize;

use xlex_core::{CellRef, Workbook};

use super::{GlobalOptions, OutputFormat};

/// Arguments for global search.
#[derive(Parser)]
pub struct SearchArgs {
    /// Path to the xlsx file
    pub file: std::path::PathBuf,

    /// Text pattern to search for
    pub pattern: String,

    /// Restrict search to a specific sheet
    #[arg(short, long)]
    pub sheet: Option<String>,

    /// Restrict search to a specific column (e.g., A, B, AA)
    #[arg(short = 'c', long)]
    pub column: Option<String>,

    /// Use case-sensitive matching (default: case-insensitive)
    #[arg(long)]
    pub case_sensitive: bool,

    /// Use regex pattern matching
    #[arg(short = 'r', long)]
    pub regex: bool,

    /// Maximum number of results to return (0 = unlimited)
    #[arg(short = 'n', long, default_value = "0")]
    pub max_results: usize,
}

/// A single search match result.
#[derive(Debug, serde::Serialize)]
struct SearchMatch {
    sheet: String,
    cell: String,
    row: u32,
    col: u32,
    value: String,
}

/// Run the search command.
pub fn run(args: &SearchArgs, global: &GlobalOptions) -> Result<()> {
    let workbook = Workbook::open(&args.file)?;

    // Build list of sheets to search
    let sheet_names: Vec<String> = if let Some(ref name) = args.sheet {
        // Verify sheet exists
        if workbook.get_sheet(name).is_none() {
            return Err(xlex_core::XlexError::SheetNotFound { name: name.clone() }.into());
        }
        vec![name.clone()]
    } else {
        workbook
            .sheet_names()
            .iter()
            .map(|s| s.to_string())
            .collect()
    };

    // Parse column filter
    let col_filter: Option<u32> = args
        .column
        .as_ref()
        .and_then(|c| CellRef::col_from_letters_pub(&c.to_uppercase()));

    // Compile regex if needed
    let compiled_regex = if args.regex {
        let pattern = if args.case_sensitive {
            regex_lite::Regex::new(&args.pattern)?
        } else {
            // regex-lite doesn't have RegexBuilder, use inline flag (?i)
            let case_insensitive_pattern = format!("(?i){}", &args.pattern);
            regex_lite::Regex::new(&case_insensitive_pattern)?
        };
        Some(pattern)
    } else {
        None
    };

    // Prepare pattern for plain text matching
    let search_pattern = if args.case_sensitive {
        args.pattern.clone()
    } else {
        args.pattern.to_lowercase()
    };

    let mut matches: Vec<SearchMatch> = Vec::new();
    let max = args.max_results;

    'outer: for sheet_name in &sheet_names {
        let sheet = match workbook.get_sheet(sheet_name) {
            Some(s) => s,
            None => continue,
        };

        for cell in sheet.cells() {
            // Apply column filter
            if let Some(col) = col_filter {
                if cell.reference.col != col {
                    continue;
                }
            }

            let value = cell.value.to_display_string();
            let is_match = if let Some(ref re) = compiled_regex {
                re.is_match(&value)
            } else if args.case_sensitive {
                value.contains(&search_pattern)
            } else {
                value.to_lowercase().contains(&search_pattern)
            };

            if is_match {
                matches.push(SearchMatch {
                    sheet: sheet_name.clone(),
                    cell: cell.reference.to_a1(),
                    row: cell.reference.row,
                    col: cell.reference.col,
                    value,
                });

                if max > 0 && matches.len() >= max {
                    break 'outer;
                }
            }
        }
    }

    // Output results
    match global.format {
        OutputFormat::Json => print_json(&matches, &args.pattern)?,
        OutputFormat::Csv => print_csv(&matches)?,
        OutputFormat::Ndjson => print_ndjson(&matches)?,
        OutputFormat::Text => print_text(&matches, &args.pattern, global)?,
    }

    Ok(())
}

fn print_text(matches: &[SearchMatch], pattern: &str, global: &GlobalOptions) -> Result<()> {
    if matches.is_empty() {
        if !global.quiet {
            println!(
                "{}: no matches found for \"{}\"",
                "search".yellow(),
                pattern
            );
        }
        return Ok(());
    }

    if !global.quiet {
        println!(
            "Found {} match{} for \"{}\":\n",
            matches.len().to_string().green(),
            if matches.len() == 1 { "" } else { "es" },
            pattern.cyan()
        );
    }

    let mut current_sheet = "";
    for m in matches {
        if m.sheet != current_sheet {
            if !current_sheet.is_empty() {
                println!();
            }
            println!("  {} {}", "Sheet:".bold(), m.sheet.bold());
            current_sheet = &m.sheet;
        }
        println!("    {} = {}", m.cell.yellow(), m.value.dimmed());
    }

    if !global.quiet {
        println!();
    }

    Ok(())
}

fn print_json(matches: &[SearchMatch], pattern: &str) -> Result<()> {
    let json = serde_json::json!({
        "pattern": pattern,
        "count": matches.len(),
        "matches": matches,
    });
    println!("{}", serde_json::to_string_pretty(&json)?);
    Ok(())
}

fn print_csv(matches: &[SearchMatch]) -> Result<()> {
    println!("sheet,cell,row,col,value");
    for m in matches {
        // Escape value for CSV
        let escaped = if m.value.contains(',') || m.value.contains('"') || m.value.contains('\n') {
            format!("\"{}\"", m.value.replace('"', "\"\""))
        } else {
            m.value.clone()
        };
        println!("{},{},{},{},{}", m.sheet, m.cell, m.row, m.col, escaped);
    }
    Ok(())
}

fn print_ndjson(matches: &[SearchMatch]) -> Result<()> {
    for m in matches {
        println!("{}", serde_json::to_string(m)?);
    }
    Ok(())
}
