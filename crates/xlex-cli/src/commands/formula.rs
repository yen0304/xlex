//! Formula operations.

use anyhow::Result;
use clap::{Parser, Subcommand};
use colored::Colorize;

use xlex_core::{CellRef, CellValue, Range, Workbook};

use super::{GlobalOptions, OutputFormat};

/// Arguments for formula operations.
#[derive(Parser)]
pub struct FormulaArgs {
    #[command(subcommand)]
    pub command: FormulaCommand,
}

#[derive(Subcommand)]
pub enum FormulaCommand {
    /// Get formula from a cell
    Get {
        /// Path to the xlsx file
        file: std::path::PathBuf,
        /// Sheet name
        sheet: String,
        /// Cell reference (e.g., A1)
        cell: String,
    },
    /// Set formula in a cell
    Set {
        /// Path to the xlsx file
        file: std::path::PathBuf,
        /// Sheet name
        sheet: String,
        /// Cell reference (e.g., A1)
        cell: String,
        /// Formula (without leading =)
        formula: String,
    },
    /// List all formulas in a sheet
    List {
        /// Path to the xlsx file
        file: std::path::PathBuf,
        /// Sheet name
        sheet: String,
    },
    /// Evaluate a formula (display result)
    Eval {
        /// Path to the xlsx file
        file: std::path::PathBuf,
        /// Sheet name
        sheet: String,
        /// Formula to evaluate
        formula: String,
    },
    /// Check formulas for errors
    Check {
        /// Path to the xlsx file
        file: std::path::PathBuf,
        /// Sheet name (check all if not specified)
        sheet: Option<String>,
    },
    /// Validate formula syntax
    Validate {
        /// Formula to validate
        formula: String,
    },
    /// Show formula statistics
    Stats {
        /// Path to the xlsx file
        file: std::path::PathBuf,
        /// Sheet name (stats for all if not specified)
        sheet: Option<String>,
    },
    /// Find formula references (dependents/precedents)
    Refs {
        /// Path to the xlsx file
        file: std::path::PathBuf,
        /// Sheet name
        sheet: String,
        /// Cell reference
        cell: String,
        /// Show dependents (cells that depend on this cell)
        #[arg(long)]
        dependents: bool,
        /// Show precedents (cells this cell depends on)
        #[arg(long)]
        precedents: bool,
    },
    /// Replace formula references
    Replace {
        /// Path to the xlsx file
        file: std::path::PathBuf,
        /// Sheet name
        sheet: String,
        /// Find pattern
        find: String,
        /// Replace with
        replace: String,
    },
    /// Calculate built-in functions
    Calc(CalcArgs),
    /// Detect circular references
    Circular {
        /// Path to the xlsx file
        file: std::path::PathBuf,
        /// Sheet name (check all if not specified)
        sheet: Option<String>,
    },
}

#[derive(Parser)]
pub struct CalcArgs {
    #[command(subcommand)]
    pub command: CalcCommand,
}

#[derive(Subcommand)]
pub enum CalcCommand {
    /// Sum values in a range
    Sum {
        /// Path to the xlsx file
        file: std::path::PathBuf,
        /// Sheet name
        sheet: String,
        /// Range (e.g., A1:A10)
        range: String,
    },
    /// Average values in a range
    Avg {
        /// Path to the xlsx file
        file: std::path::PathBuf,
        /// Sheet name
        sheet: String,
        /// Range
        range: String,
    },
    /// Count values in a range
    Count {
        /// Path to the xlsx file
        file: std::path::PathBuf,
        /// Sheet name
        sheet: String,
        /// Range
        range: String,
        /// Count only non-empty cells
        #[arg(long)]
        nonempty: bool,
    },
    /// Get minimum value in a range
    Min {
        /// Path to the xlsx file
        file: std::path::PathBuf,
        /// Sheet name
        sheet: String,
        /// Range
        range: String,
    },
    /// Get maximum value in a range
    Max {
        /// Path to the xlsx file
        file: std::path::PathBuf,
        /// Sheet name
        sheet: String,
        /// Range
        range: String,
    },
}

/// Run formula operations.
pub fn run(args: &FormulaArgs, global: &GlobalOptions) -> Result<()> {
    match &args.command {
        FormulaCommand::Get { file, sheet, cell } => get(file, sheet, cell, global),
        FormulaCommand::Set {
            file,
            sheet,
            cell,
            formula,
        } => set(file, sheet, cell, formula, global),
        FormulaCommand::List { file, sheet } => list(file, sheet, global),
        FormulaCommand::Eval {
            file,
            sheet,
            formula,
        } => eval(file, sheet, formula, global),
        FormulaCommand::Check { file, sheet } => check(file, sheet.as_deref(), global),
        FormulaCommand::Validate { formula } => validate(formula, global),
        FormulaCommand::Stats { file, sheet } => stats(file, sheet.as_deref(), global),
        FormulaCommand::Refs {
            file,
            sheet,
            cell,
            dependents,
            precedents,
        } => refs(file, sheet, cell, *dependents, *precedents, global),
        FormulaCommand::Replace {
            file,
            sheet,
            find,
            replace,
        } => replace_formula(file, sheet, find, replace, global),
        FormulaCommand::Calc(calc_args) => run_calc(calc_args, global),
        FormulaCommand::Circular { file, sheet } => circular(file, sheet.as_deref(), global),
    }
}

fn get(file: &std::path::Path, sheet: &str, cell: &str, global: &GlobalOptions) -> Result<()> {
    let workbook = Workbook::open(file)?;
    let cell_ref = CellRef::parse(cell)?;
    let value = workbook.get_cell(sheet, &cell_ref)?;

    let formula = match &value {
        CellValue::Formula { formula: f, .. } => Some(f.clone()),
        _ => None,
    };

    if global.format == OutputFormat::Json {
        let json = serde_json::json!({
            "cell": cell,
            "formula": formula,
            "value": value.to_display_string(),
        });
        println!("{}", serde_json::to_string_pretty(&json)?);
    } else if let Some(f) = formula {
        println!("={}", f);
    } else {
        println!("(no formula)");
    }

    Ok(())
}

fn set(
    file: &std::path::Path,
    sheet: &str,
    cell: &str,
    formula: &str,
    global: &GlobalOptions,
) -> Result<()> {
    if global.dry_run {
        println!("Would set formula ={} in {}", formula, cell);
        return Ok(());
    }

    let mut workbook = Workbook::open(file)?;
    let cell_ref = CellRef::parse(cell)?;

    // Store formula (result is None - will be calculated by Excel)
    workbook.set_cell(sheet, cell_ref.clone(), CellValue::formula(formula))?;

    workbook.save()?;

    if !global.quiet {
        println!(
            "Set formula {} = {}",
            cell.cyan(),
            format!("={}", formula).green()
        );
    }

    Ok(())
}

fn list(file: &std::path::Path, sheet: &str, global: &GlobalOptions) -> Result<()> {
    let workbook = Workbook::open(file)?;
    let sheet_obj =
        workbook
            .get_sheet(sheet)
            .ok_or_else(|| xlex_core::XlexError::SheetNotFound {
                name: sheet.to_string(),
            })?;

    let formulas: Vec<_> = sheet_obj
        .cells()
        .filter_map(|c| match &c.value {
            CellValue::Formula {
                formula: f,
                cached_result: result,
            } => Some((c.reference.clone(), f.clone(), result.clone())),
            _ => None,
        })
        .collect();

    if global.format == OutputFormat::Json {
        let json: Vec<_> = formulas
            .iter()
            .map(|(cell, formula, result)| {
                serde_json::json!({
                    "cell": cell.to_a1(),
                    "formula": formula,
                    "result": result.as_ref().map(|r| r.to_display_string()),
                })
            })
            .collect();
        println!("{}", serde_json::to_string_pretty(&json)?);
    } else {
        println!("{}: {} formulas", "Sheet".bold(), formulas.len());
        for (cell, formula, result) in &formulas {
            let result_str = result
                .as_ref()
                .map(|r| format!(" → {}", r.to_display_string()))
                .unwrap_or_default();
            println!(
                "  {}: ={}{}",
                cell.to_a1().cyan(),
                formula,
                result_str.dimmed()
            );
        }
    }

    Ok(())
}

fn eval(
    _file: &std::path::Path,
    _sheet: &str,
    formula: &str,
    global: &GlobalOptions,
) -> Result<()> {
    // Basic formula evaluation (very limited - Excel does the real work)
    // This is just for simple expressions

    let result = if formula.starts_with('=') {
        &formula[1..]
    } else {
        formula
    };

    // Try to parse as simple number expression
    if let Ok(n) = result.parse::<f64>() {
        if global.format == OutputFormat::Json {
            println!("{}", serde_json::json!({ "formula": formula, "result": n }));
        } else {
            println!("{}", n);
        }
        return Ok(());
    }

    // For complex formulas, we can't evaluate without Excel
    println!("Formula evaluation requires Excel - result will be calculated when file is opened");
    Ok(())
}

fn check(file: &std::path::Path, sheet: Option<&str>, global: &GlobalOptions) -> Result<()> {
    let workbook = Workbook::open(file)?;

    let sheets: Vec<_> = if let Some(s) = sheet {
        vec![s.to_string()]
    } else {
        workbook
            .sheet_names()
            .iter()
            .map(|s| s.to_string())
            .collect::<Vec<String>>()
    };

    let mut errors: Vec<(String, String, String)> = Vec::new();

    for sheet_name in &sheets {
        if let Some(sheet_obj) = workbook.get_sheet(sheet_name) {
            for cell in sheet_obj.cells() {
                if let CellValue::Formula {
                    cached_result: Some(result),
                    ..
                } = &cell.value
                {
                    if matches!(result.as_ref(), CellValue::Error(_)) {
                        errors.push((
                            sheet_name.to_string(),
                            cell.reference.to_a1(),
                            result.to_display_string(),
                        ));
                    }
                }
            }
        }
    }

    if global.format == OutputFormat::Json {
        let json = serde_json::json!({
            "errors": errors.iter().map(|(s, c, e)| {
                serde_json::json!({
                    "sheet": s,
                    "cell": c,
                    "error": e,
                })
            }).collect::<Vec<_>>(),
            "count": errors.len(),
        });
        println!("{}", serde_json::to_string_pretty(&json)?);
    } else if errors.is_empty() {
        println!("{}: No formula errors found", "✓".green());
    } else {
        println!("{}: {} formula errors found", "✗".red(), errors.len());
        for (sheet_name, cell, error) in &errors {
            println!("  {}!{}: {}", sheet_name, cell.cyan(), error.red());
        }
    }

    Ok(())
}

fn run_calc(args: &CalcArgs, global: &GlobalOptions) -> Result<()> {
    match &args.command {
        CalcCommand::Sum { file, sheet, range } => calc_sum(file, sheet, range, global),
        CalcCommand::Avg { file, sheet, range } => calc_avg(file, sheet, range, global),
        CalcCommand::Count {
            file,
            sheet,
            range,
            nonempty,
        } => calc_count(file, sheet, range, *nonempty, global),
        CalcCommand::Min { file, sheet, range } => calc_min(file, sheet, range, global),
        CalcCommand::Max { file, sheet, range } => calc_max(file, sheet, range, global),
    }
}

fn calc_sum(
    file: &std::path::Path,
    sheet: &str,
    range: &str,
    global: &GlobalOptions,
) -> Result<()> {
    let workbook = Workbook::open(file)?;
    let range_ref = Range::parse(range)?;
    let sheet_obj =
        workbook
            .get_sheet(sheet)
            .ok_or_else(|| xlex_core::XlexError::SheetNotFound {
                name: sheet.to_string(),
            })?;

    let mut sum = 0.0;
    for cell_ref in range_ref.cells() {
        let value = sheet_obj.get_value(&cell_ref);
        if let CellValue::Number(n) = value {
            sum += n;
        }
    }

    if global.format == OutputFormat::Json {
        println!("{}", serde_json::json!({ "sum": sum }));
    } else {
        println!("{}", sum);
    }

    Ok(())
}

fn calc_avg(
    file: &std::path::Path,
    sheet: &str,
    range: &str,
    global: &GlobalOptions,
) -> Result<()> {
    let workbook = Workbook::open(file)?;
    let range_ref = Range::parse(range)?;
    let sheet_obj =
        workbook
            .get_sheet(sheet)
            .ok_or_else(|| xlex_core::XlexError::SheetNotFound {
                name: sheet.to_string(),
            })?;

    let mut sum = 0.0;
    let mut count = 0;
    for cell_ref in range_ref.cells() {
        let value = sheet_obj.get_value(&cell_ref);
        if let CellValue::Number(n) = value {
            sum += n;
            count += 1;
        }
    }

    let avg = if count > 0 { sum / count as f64 } else { 0.0 };

    if global.format == OutputFormat::Json {
        println!("{}", serde_json::json!({ "average": avg, "count": count }));
    } else {
        println!("{}", avg);
    }

    Ok(())
}

fn calc_count(
    file: &std::path::Path,
    sheet: &str,
    range: &str,
    nonempty: bool,
    global: &GlobalOptions,
) -> Result<()> {
    let workbook = Workbook::open(file)?;
    let range_ref = Range::parse(range)?;
    let sheet_obj =
        workbook
            .get_sheet(sheet)
            .ok_or_else(|| xlex_core::XlexError::SheetNotFound {
                name: sheet.to_string(),
            })?;

    let mut count = 0;
    for cell_ref in range_ref.cells() {
        let value = sheet_obj.get_value(&cell_ref);
        if nonempty {
            if !value.is_empty() {
                count += 1;
            }
        } else {
            count += 1;
        }
    }

    if global.format == OutputFormat::Json {
        println!("{}", serde_json::json!({ "count": count }));
    } else {
        println!("{}", count);
    }

    Ok(())
}

fn calc_min(
    file: &std::path::Path,
    sheet: &str,
    range: &str,
    global: &GlobalOptions,
) -> Result<()> {
    let workbook = Workbook::open(file)?;
    let range_ref = Range::parse(range)?;
    let sheet_obj =
        workbook
            .get_sheet(sheet)
            .ok_or_else(|| xlex_core::XlexError::SheetNotFound {
                name: sheet.to_string(),
            })?;

    let mut min: Option<f64> = None;
    for cell_ref in range_ref.cells() {
        let value = sheet_obj.get_value(&cell_ref);
        if let CellValue::Number(n) = value {
            min = Some(min.map_or(n, |m| m.min(n)));
        }
    }

    if global.format == OutputFormat::Json {
        println!("{}", serde_json::json!({ "min": min }));
    } else {
        match min {
            Some(m) => println!("{}", m),
            None => println!("(no numeric values)"),
        }
    }

    Ok(())
}

fn calc_max(
    file: &std::path::Path,
    sheet: &str,
    range: &str,
    global: &GlobalOptions,
) -> Result<()> {
    let workbook = Workbook::open(file)?;
    let range_ref = Range::parse(range)?;
    let sheet_obj =
        workbook
            .get_sheet(sheet)
            .ok_or_else(|| xlex_core::XlexError::SheetNotFound {
                name: sheet.to_string(),
            })?;

    let mut max: Option<f64> = None;
    for cell_ref in range_ref.cells() {
        let value = sheet_obj.get_value(&cell_ref);
        if let CellValue::Number(n) = value {
            max = Some(max.map_or(n, |m| m.max(n)));
        }
    }

    if global.format == OutputFormat::Json {
        println!("{}", serde_json::json!({ "max": max }));
    } else {
        match max {
            Some(m) => println!("{}", m),
            None => println!("(no numeric values)"),
        }
    }

    Ok(())
}

fn validate(formula: &str, global: &GlobalOptions) -> Result<()> {
    // Basic formula validation - check for common patterns
    let formula_str = if formula.starts_with('=') {
        &formula[1..]
    } else {
        formula
    };

    let mut errors: Vec<String> = Vec::new();
    let mut warnings: Vec<String> = Vec::new();

    // Check for balanced parentheses
    let mut paren_count = 0;
    for (i, c) in formula_str.chars().enumerate() {
        match c {
            '(' => paren_count += 1,
            ')' => {
                paren_count -= 1;
                if paren_count < 0 {
                    errors.push(format!("Unmatched ')' at position {}", i));
                }
            }
            _ => {}
        }
    }
    if paren_count > 0 {
        errors.push(format!("{} unclosed '('", paren_count));
    }

    // Check for common syntax issues
    if formula_str.contains(",,") {
        warnings.push("Empty argument detected (double comma)".to_string());
    }
    if formula_str.ends_with(',') || formula_str.ends_with('(') {
        errors.push("Formula ends with incomplete expression".to_string());
    }

    // Extract function names and check if they look valid
    let re_func = regex_lite::Regex::new(r"[A-Z]+\s*\(").ok();
    if let Some(re) = re_func {
        for cap in re.find_iter(formula_str) {
            let func_name = cap.as_str().trim_end_matches(&['(', ' '][..]);
            // Known Excel functions (sample list)
            let known_funcs = [
                "SUM",
                "AVERAGE",
                "COUNT",
                "COUNTA",
                "MIN",
                "MAX",
                "IF",
                "VLOOKUP",
                "HLOOKUP",
                "INDEX",
                "MATCH",
                "SUMIF",
                "COUNTIF",
                "CONCATENATE",
                "LEFT",
                "RIGHT",
                "MID",
                "LEN",
                "TRIM",
                "UPPER",
                "LOWER",
                "ROUND",
                "ABS",
                "SQRT",
                "NOW",
                "TODAY",
                "DATE",
                "YEAR",
                "MONTH",
                "DAY",
                "AND",
                "OR",
                "NOT",
                "TRUE",
                "FALSE",
            ];
            if !known_funcs
                .iter()
                .any(|f| f.eq_ignore_ascii_case(func_name))
            {
                warnings.push(format!("Unknown function: {}", func_name));
            }
        }
    }

    let is_valid = errors.is_empty();

    if global.format == OutputFormat::Json {
        let json = serde_json::json!({
            "formula": formula,
            "valid": is_valid,
            "errors": errors,
            "warnings": warnings,
        });
        println!("{}", serde_json::to_string_pretty(&json)?);
    } else {
        if is_valid {
            println!("{}: Formula syntax is valid", "✓".green());
        } else {
            println!("{}: Formula has syntax errors", "✗".red());
        }
        for err in &errors {
            println!("  {}: {}", "Error".red(), err);
        }
        for warn in &warnings {
            println!("  {}: {}", "Warning".yellow(), warn);
        }
    }

    if is_valid {
        Ok(())
    } else {
        anyhow::bail!("Formula validation failed")
    }
}

fn stats(file: &std::path::Path, sheet: Option<&str>, global: &GlobalOptions) -> Result<()> {
    let workbook = Workbook::open(file)?;

    let sheets: Vec<_> = if let Some(s) = sheet {
        vec![s.to_string()]
    } else {
        workbook
            .sheet_names()
            .iter()
            .map(|s| s.to_string())
            .collect::<Vec<String>>()
    };

    let mut total_formulas = 0;
    let mut total_errors = 0;
    let mut function_counts: std::collections::HashMap<String, usize> =
        std::collections::HashMap::new();

    for sheet_name in &sheets {
        if let Some(sheet_obj) = workbook.get_sheet(sheet_name) {
            for cell in sheet_obj.cells() {
                if let CellValue::Formula {
                    formula: f,
                    cached_result: result,
                } = &cell.value
                {
                    total_formulas += 1;

                    // Check for errors
                    if let Some(r) = result {
                        if matches!(r.as_ref(), CellValue::Error(_)) {
                            total_errors += 1;
                        }
                    }

                    // Extract function names
                    let re = regex_lite::Regex::new(r"[A-Z]+\s*\(").ok();
                    if let Some(re) = re {
                        for cap in re.find_iter(f) {
                            let func_name = cap
                                .as_str()
                                .trim_end_matches(&['(', ' '][..])
                                .to_uppercase();
                            *function_counts.entry(func_name).or_insert(0) += 1;
                        }
                    }
                }
            }
        }
    }

    // Sort functions by count
    let mut sorted_funcs: Vec<_> = function_counts.into_iter().collect();
    sorted_funcs.sort_by(|a, b| b.1.cmp(&a.1));

    if global.format == OutputFormat::Json {
        let json = serde_json::json!({
            "totalFormulas": total_formulas,
            "errorCount": total_errors,
            "functions": sorted_funcs.iter().map(|(name, count)| {
                serde_json::json!({"name": name, "count": count})
            }).collect::<Vec<_>>(),
        });
        println!("{}", serde_json::to_string_pretty(&json)?);
    } else {
        println!("{}: {}", "Total formulas".bold(), total_formulas);
        println!("{}: {}", "Errors".bold(), total_errors);
        println!("\n{}:", "Functions used".bold());
        for (func, count) in sorted_funcs.iter().take(10) {
            println!("  {}: {}", func.cyan(), count);
        }
        if sorted_funcs.len() > 10 {
            println!("  ... and {} more", sorted_funcs.len() - 10);
        }
    }

    Ok(())
}

fn refs(
    file: &std::path::Path,
    sheet: &str,
    cell: &str,
    show_dependents: bool,
    show_precedents: bool,
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

    let mut precedents: Vec<String> = Vec::new();
    let mut dependents: Vec<String> = Vec::new();

    // Get precedents (cells referenced by this cell's formula)
    if show_precedents || (!show_dependents && !show_precedents) {
        let value = sheet_obj.get_value(&cell_ref);
        if let CellValue::Formula { formula: f, .. } = value {
            // Extract cell references from formula
            let re = regex_lite::Regex::new(r"\$?[A-Z]+\$?\d+").ok();
            if let Some(re) = re {
                for cap in re.find_iter(&f) {
                    precedents.push(cap.as_str().to_string());
                }
            }
        }
    }

    // Get dependents (cells that reference this cell)
    if show_dependents || (!show_dependents && !show_precedents) {
        let cell_a1 = cell_ref.to_a1();
        for c in sheet_obj.cells() {
            if let CellValue::Formula { formula: f, .. } = &c.value {
                // Check if formula references this cell
                if f.contains(&cell_a1) || f.contains(&format!("${}", cell_a1)) {
                    dependents.push(c.reference.to_a1());
                }
            }
        }
    }

    if global.format == OutputFormat::Json {
        let json = serde_json::json!({
            "cell": cell,
            "precedents": precedents,
            "dependents": dependents,
        });
        println!("{}", serde_json::to_string_pretty(&json)?);
    } else {
        if !precedents.is_empty() {
            println!("{} (cells referenced by {}):", "Precedents".bold(), cell);
            for p in &precedents {
                println!("  {}", p.cyan());
            }
        }
        if !dependents.is_empty() {
            println!("{} (cells that reference {}):", "Dependents".bold(), cell);
            for d in &dependents {
                println!("  {}", d.cyan());
            }
        }
        if precedents.is_empty() && dependents.is_empty() {
            println!("No references found for {}", cell);
        }
    }

    Ok(())
}

fn replace_formula(
    file: &std::path::Path,
    sheet: &str,
    find: &str,
    replace_with: &str,
    global: &GlobalOptions,
) -> Result<()> {
    if global.dry_run {
        println!(
            "Would replace '{}' with '{}' in formulas",
            find, replace_with
        );
        return Ok(());
    }

    let mut workbook = Workbook::open(file)?;
    let sheet_obj =
        workbook
            .get_sheet_mut(sheet)
            .ok_or_else(|| xlex_core::XlexError::SheetNotFound {
                name: sheet.to_string(),
            })?;

    let mut replaced_count = 0;

    // Collect cells with formulas that need updating
    let cells_to_update: Vec<_> = sheet_obj
        .cells()
        .filter_map(|c| {
            if let CellValue::Formula { formula: f, .. } = &c.value {
                if f.contains(find) {
                    Some((c.reference.clone(), f.clone()))
                } else {
                    None
                }
            } else {
                None
            }
        })
        .collect();

    // Update formulas
    for (cell_ref, formula) in cells_to_update {
        let new_formula = formula.replace(find, replace_with);
        sheet_obj.set_cell(cell_ref, CellValue::formula(&new_formula));
        replaced_count += 1;
    }

    let _ = sheet_obj;
    workbook.save()?;

    if global.format == OutputFormat::Json {
        let json = serde_json::json!({
            "find": find,
            "replace": replace_with,
            "count": replaced_count,
        });
        println!("{}", serde_json::to_string_pretty(&json)?);
    } else if !global.quiet {
        println!(
            "Replaced '{}' with '{}' in {} formulas",
            find.cyan(),
            replace_with.green(),
            replaced_count
        );
    }

    Ok(())
}
/// Detect circular references in formulas.
fn circular(
    file: &std::path::Path,
    sheet_name: Option<&str>,
    global: &GlobalOptions,
) -> Result<()> {
    use std::collections::{HashMap, HashSet};

    let workbook = Workbook::open(file)?;

    // Collect all formulas and their references
    let sheets_to_check: Vec<String> = if let Some(name) = sheet_name {
        vec![name.to_string()]
    } else {
        workbook
            .sheet_names()
            .iter()
            .map(|s| s.to_string())
            .collect()
    };

    let mut all_circular: Vec<(String, String, Vec<String>)> = Vec::new(); // (sheet, cell, cycle_path)

    for sheet_name in &sheets_to_check {
        let Some(sheet) = workbook.get_sheet(sheet_name) else {
            continue;
        };

        // Build dependency graph: cell -> cells it depends on
        let mut deps: HashMap<String, Vec<String>> = HashMap::new();

        for cell in sheet.cells() {
            if let CellValue::Formula { formula, .. } = &cell.value {
                let cell_key = format!("{}!{}", sheet_name, cell.reference.to_a1());
                let refs = extract_cell_refs(formula);
                let qualified_refs: Vec<String> = refs
                    .into_iter()
                    .map(|r| {
                        if r.contains('!') {
                            r
                        } else {
                            format!("{}!{}", sheet_name, r)
                        }
                    })
                    .collect();
                deps.insert(cell_key, qualified_refs);
            }
        }

        // Detect cycles using DFS
        for start_cell in deps.keys() {
            let mut visited = HashSet::new();
            let mut path = Vec::new();
            if let Some(cycle) = detect_cycle(&deps, start_cell, &mut visited, &mut path) {
                // Only report if this cell is part of the cycle
                if cycle.contains(start_cell) {
                    let cell_only = start_cell.split('!').last().unwrap_or(start_cell);
                    all_circular.push((sheet_name.to_string(), cell_only.to_string(), cycle));
                }
            }
        }
    }

    // Deduplicate cycles (same cycle may be detected from different starting points)
    let mut unique_cycles: Vec<(String, String, Vec<String>)> = Vec::new();
    let mut seen_cycles: HashSet<String> = HashSet::new();
    for (sheet, cell, cycle) in all_circular {
        let mut sorted_cycle = cycle.clone();
        sorted_cycle.sort();
        let cycle_key = sorted_cycle.join(",");
        if !seen_cycles.contains(&cycle_key) {
            seen_cycles.insert(cycle_key);
            unique_cycles.push((sheet, cell, cycle));
        }
    }

    if global.format == OutputFormat::Json {
        let json: Vec<_> = unique_cycles
            .iter()
            .map(|(sheet, cell, cycle)| {
                serde_json::json!({
                    "sheet": sheet,
                    "cell": cell,
                    "cycle": cycle,
                })
            })
            .collect();
        println!("{}", serde_json::to_string_pretty(&json)?);
    } else if unique_cycles.is_empty() {
        if !global.quiet {
            println!("{}", "No circular references detected".green());
        }
    } else {
        println!(
            "{}: {} circular reference(s) detected",
            "Warning".yellow().bold(),
            unique_cycles.len()
        );
        for (sheet, cell, cycle) in &unique_cycles {
            println!(
                "\n  {} in {}!{}:",
                "Cycle".red().bold(),
                sheet.cyan(),
                cell.cyan()
            );
            println!("    {}", cycle.join(" → "));
        }
    }

    Ok(())
}

/// Extract cell references from a formula.
fn extract_cell_refs(formula: &str) -> Vec<String> {
    use regex_lite::Regex;

    // Match cell references like A1, $A$1, Sheet1!A1, 'Sheet Name'!A1
    // Also match ranges like A1:B10
    let re = Regex::new(r"(?:'[^']+'!)?(?:\$?[A-Z]+\$?\d+)(?::\$?[A-Z]+\$?\d+)?").unwrap();

    let mut refs = Vec::new();
    for cap in re.find_iter(formula) {
        let cell_ref = cap.as_str().to_string();
        // If it's a range, expand to individual cells (simplified - just add start and end)
        if cell_ref.contains(':') {
            let parts: Vec<&str> = cell_ref.split(':').collect();
            if parts.len() == 2 {
                refs.push(parts[0].to_string());
                refs.push(parts[1].to_string());
            }
        } else {
            refs.push(cell_ref);
        }
    }

    refs
}

/// Detect cycle in dependency graph using DFS.
fn detect_cycle(
    deps: &std::collections::HashMap<String, Vec<String>>,
    current: &str,
    visited: &mut std::collections::HashSet<String>,
    path: &mut Vec<String>,
) -> Option<Vec<String>> {
    if path.contains(&current.to_string()) {
        // Found a cycle - return the cycle portion
        let cycle_start = path.iter().position(|x| x == current).unwrap();
        let mut cycle: Vec<String> = path[cycle_start..].to_vec();
        cycle.push(current.to_string());
        return Some(cycle);
    }

    if visited.contains(current) {
        return None;
    }

    visited.insert(current.to_string());
    path.push(current.to_string());

    if let Some(dependencies) = deps.get(current) {
        for dep in dependencies {
            if let Some(cycle) = detect_cycle(deps, dep, visited, path) {
                return Some(cycle);
            }
        }
    }

    path.pop();
    None
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

    fn setup_formula_data(file: &std::path::Path) {
        let mut wb = Workbook::open(file).unwrap();
        // Set some values
        wb.set_cell("Sheet1", CellRef::new(1, 1), CellValue::Number(10.0))
            .unwrap();
        wb.set_cell("Sheet1", CellRef::new(2, 1), CellValue::Number(20.0))
            .unwrap();
        wb.set_cell("Sheet1", CellRef::new(3, 1), CellValue::Number(30.0))
            .unwrap();
        // Set a formula
        wb.set_cell(
            "Sheet1",
            CellRef::new(1, 2),
            CellValue::formula("SUM(A1:C1)"),
        )
        .unwrap();
        wb.save().unwrap();
    }

    #[test]
    fn test_get_formula() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "get.xlsx");
        setup_formula_data(&file_path);

        let result = get(&file_path, "Sheet1", "A2", &default_global());
        assert!(result.is_ok());
    }

    #[test]
    fn test_get_non_formula() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "get_non.xlsx");
        setup_formula_data(&file_path);

        let result = get(&file_path, "Sheet1", "A1", &default_global());
        assert!(result.is_ok());
    }

    #[test]
    fn test_set_formula() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "set.xlsx");

        let result = set(&file_path, "Sheet1", "A1", "SUM(B1:B10)", &default_global());
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
    fn test_set_formula_dry_run() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "set_dry.xlsx");

        let mut global = default_global();
        global.dry_run = true;

        let result = set(&file_path, "Sheet1", "A1", "SUM(B1:B10)", &global);
        assert!(result.is_ok());
    }

    #[test]
    fn test_list_formulas() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "list.xlsx");
        setup_formula_data(&file_path);

        let result = list(&file_path, "Sheet1", &default_global());
        assert!(result.is_ok());
    }

    #[test]
    fn test_check_formulas() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "check.xlsx");
        setup_formula_data(&file_path);

        let result = check(&file_path, Some("Sheet1"), &default_global());
        assert!(result.is_ok());
    }

    #[test]
    fn test_check_all_sheets() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "check_all.xlsx");
        setup_formula_data(&file_path);

        let result = check(&file_path, None, &default_global());
        assert!(result.is_ok());
    }

    #[test]
    fn test_validate_valid_formula() {
        let result = validate("SUM(A1:A10)", &default_global());
        assert!(result.is_ok());
    }

    #[test]
    fn test_validate_with_equals() {
        let result = validate("=SUM(A1:A10)", &default_global());
        assert!(result.is_ok());
    }

    #[test]
    fn test_validate_invalid_formula() {
        let result = validate("SUM(A1:A10", &default_global());
        assert!(result.is_err()); // Missing closing paren
    }

    #[test]
    fn test_stats() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "stats.xlsx");
        setup_formula_data(&file_path);

        let result = stats(&file_path, Some("Sheet1"), &default_global());
        assert!(result.is_ok());
    }

    #[test]
    fn test_refs() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "refs.xlsx");
        setup_formula_data(&file_path);

        let result = refs(&file_path, "Sheet1", "A2", false, true, &default_global());
        assert!(result.is_ok());
    }

    #[test]
    fn test_replace_formula() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "replace.xlsx");
        setup_formula_data(&file_path);

        let result = replace_formula(&file_path, "Sheet1", "A1", "A2", &default_global());
        assert!(result.is_ok());
    }

    #[test]
    fn test_calc_sum() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "sum.xlsx");
        setup_formula_data(&file_path);

        let result = calc_sum(&file_path, "Sheet1", "A1:C1", &default_global());
        assert!(result.is_ok());
    }

    #[test]
    fn test_calc_avg() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "avg.xlsx");
        setup_formula_data(&file_path);

        let result = calc_avg(&file_path, "Sheet1", "A1:C1", &default_global());
        assert!(result.is_ok());
    }

    #[test]
    fn test_calc_count() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "count.xlsx");
        setup_formula_data(&file_path);

        let result = calc_count(&file_path, "Sheet1", "A1:C1", false, &default_global());
        assert!(result.is_ok());
    }

    #[test]
    fn test_calc_count_nonempty() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "count_ne.xlsx");
        setup_formula_data(&file_path);

        let result = calc_count(&file_path, "Sheet1", "A1:C1", true, &default_global());
        assert!(result.is_ok());
    }

    #[test]
    fn test_calc_min() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "min.xlsx");
        setup_formula_data(&file_path);

        let result = calc_min(&file_path, "Sheet1", "A1:C1", &default_global());
        assert!(result.is_ok());
    }

    #[test]
    fn test_calc_max() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "max.xlsx");
        setup_formula_data(&file_path);

        let result = calc_max(&file_path, "Sheet1", "A1:C1", &default_global());
        assert!(result.is_ok());
    }

    #[test]
    fn test_circular_detection() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "circular.xlsx");

        let result = circular(&file_path, Some("Sheet1"), &default_global());
        assert!(result.is_ok());
    }

    #[test]
    fn test_extract_cell_refs() {
        let refs = extract_cell_refs("SUM(A1:B10)");
        assert!(refs.contains(&"A1".to_string()));
        assert!(refs.contains(&"B10".to_string()));
    }

    #[test]
    fn test_extract_cell_refs_single() {
        let refs = extract_cell_refs("A1+B1");
        assert!(refs.contains(&"A1".to_string()));
        assert!(refs.contains(&"B1".to_string()));
    }

    // Additional tests for better coverage

    #[test]
    fn test_run_get_command() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "run_get.xlsx");
        setup_formula_data(&file_path);

        let args = FormulaArgs {
            command: FormulaCommand::Get {
                file: file_path,
                sheet: "Sheet1".to_string(),
                cell: "A2".to_string(),
            },
        };

        let result = run(&args, &default_global());
        assert!(result.is_ok());
    }

    #[test]
    fn test_run_set_command() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "run_set.xlsx");

        let args = FormulaArgs {
            command: FormulaCommand::Set {
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
    fn test_run_list_command() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "run_list.xlsx");
        setup_formula_data(&file_path);

        let args = FormulaArgs {
            command: FormulaCommand::List {
                file: file_path,
                sheet: "Sheet1".to_string(),
            },
        };

        let result = run(&args, &default_global());
        assert!(result.is_ok());
    }

    #[test]
    fn test_run_check_command() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "run_check.xlsx");
        setup_formula_data(&file_path);

        let args = FormulaArgs {
            command: FormulaCommand::Check {
                file: file_path,
                sheet: Some("Sheet1".to_string()),
            },
        };

        let result = run(&args, &default_global());
        assert!(result.is_ok());
    }

    #[test]
    fn test_run_validate_command() {
        let args = FormulaArgs {
            command: FormulaCommand::Validate {
                formula: "SUM(A1:A10)".to_string(),
            },
        };

        let result = run(&args, &default_global());
        assert!(result.is_ok());
    }

    #[test]
    fn test_run_stats_command() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "run_stats.xlsx");
        setup_formula_data(&file_path);

        let args = FormulaArgs {
            command: FormulaCommand::Stats {
                file: file_path,
                sheet: Some("Sheet1".to_string()),
            },
        };

        let result = run(&args, &default_global());
        assert!(result.is_ok());
    }

    #[test]
    fn test_run_refs_command() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "run_refs.xlsx");
        setup_formula_data(&file_path);

        let args = FormulaArgs {
            command: FormulaCommand::Refs {
                file: file_path,
                sheet: "Sheet1".to_string(),
                cell: "A2".to_string(),
                dependents: false,
                precedents: true,
            },
        };

        let result = run(&args, &default_global());
        assert!(result.is_ok());
    }

    #[test]
    fn test_run_replace_command() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "run_replace.xlsx");
        setup_formula_data(&file_path);

        let args = FormulaArgs {
            command: FormulaCommand::Replace {
                file: file_path,
                sheet: "Sheet1".to_string(),
                find: "A1".to_string(),
                replace: "A2".to_string(),
            },
        };

        let result = run(&args, &default_global());
        assert!(result.is_ok());
    }

    #[test]
    fn test_run_circular_command() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "run_circular.xlsx");
        setup_formula_data(&file_path);

        let args = FormulaArgs {
            command: FormulaCommand::Circular {
                file: file_path,
                sheet: Some("Sheet1".to_string()),
            },
        };

        let result = run(&args, &default_global());
        assert!(result.is_ok());
    }

    #[test]
    fn test_get_formula_json() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "get_json.xlsx");
        setup_formula_data(&file_path);

        let mut global = default_global();
        global.format = OutputFormat::Json;

        let result = get(&file_path, "Sheet1", "A2", &global);
        assert!(result.is_ok());
    }

    #[test]
    fn test_list_formulas_json() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "list_json.xlsx");
        setup_formula_data(&file_path);

        let mut global = default_global();
        global.format = OutputFormat::Json;

        let result = list(&file_path, "Sheet1", &global);
        assert!(result.is_ok());
    }

    #[test]
    fn test_stats_json() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "stats_json.xlsx");
        setup_formula_data(&file_path);

        let mut global = default_global();
        global.format = OutputFormat::Json;

        let result = stats(&file_path, Some("Sheet1"), &global);
        assert!(result.is_ok());
    }

    #[test]
    fn test_refs_dependents() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "refs_dep.xlsx");
        setup_formula_data(&file_path);

        let result = refs(&file_path, "Sheet1", "A1", true, false, &default_global());
        assert!(result.is_ok());
    }

    #[test]
    fn test_refs_json() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "refs_json.xlsx");
        setup_formula_data(&file_path);

        let mut global = default_global();
        global.format = OutputFormat::Json;

        let result = refs(&file_path, "Sheet1", "A2", false, true, &global);
        assert!(result.is_ok());
    }

    #[test]
    fn test_calc_sum_json() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "sum_json.xlsx");
        setup_formula_data(&file_path);

        let mut global = default_global();
        global.format = OutputFormat::Json;

        let result = calc_sum(&file_path, "Sheet1", "A1:C1", &global);
        assert!(result.is_ok());
    }

    #[test]
    fn test_check_json() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "check_json.xlsx");
        setup_formula_data(&file_path);

        let mut global = default_global();
        global.format = OutputFormat::Json;

        let result = check(&file_path, Some("Sheet1"), &global);
        assert!(result.is_ok());
    }

    #[test]
    fn test_validate_json() {
        let mut global = default_global();
        global.format = OutputFormat::Json;

        let result = validate("SUM(A1:A10)", &global);
        assert!(result.is_ok());
    }

    #[test]
    fn test_circular_json() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "circular_json.xlsx");
        setup_formula_data(&file_path);

        let mut global = default_global();
        global.format = OutputFormat::Json;

        let result = circular(&file_path, Some("Sheet1"), &global);
        assert!(result.is_ok());
    }

    #[test]
    fn test_replace_formula_dry_run() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "replace_dry.xlsx");
        setup_formula_data(&file_path);

        let mut global = default_global();
        global.dry_run = true;

        let result = replace_formula(&file_path, "Sheet1", "A1", "A2", &global);
        assert!(result.is_ok());
    }

    #[test]
    fn test_set_formula_json() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "set_json.xlsx");

        let mut global = default_global();
        global.format = OutputFormat::Json;
        global.quiet = false;

        let result = set(&file_path, "Sheet1", "A1", "SUM(B1:B10)", &global);
        assert!(result.is_ok());
    }

    #[test]
    fn test_stats_all_sheets() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "stats_all.xlsx");
        setup_formula_data(&file_path);

        let result = stats(&file_path, None, &default_global());
        assert!(result.is_ok());
    }

    #[test]
    fn test_circular_all_sheets() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "circular_all.xlsx");
        setup_formula_data(&file_path);

        let result = circular(&file_path, None, &default_global());
        assert!(result.is_ok());
    }
}
