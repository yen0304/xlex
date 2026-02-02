//! Template operations.

use anyhow::Result;
use clap::{Parser, Subcommand};
use colored::Colorize;
use std::collections::HashMap;

use xlex_core::{CellRef, CellValue, Workbook};

use super::{GlobalOptions, OutputFormat};

/// Arguments for template operations.
#[derive(Parser)]
pub struct TemplateArgs {
    #[command(subcommand)]
    pub command: TemplateCommand,
}

#[derive(Subcommand)]
pub enum TemplateCommand {
    /// Apply a template file with variable substitution
    Apply {
        /// Template xlsx file
        template: std::path::PathBuf,
        /// Output file
        output: std::path::PathBuf,
        /// Variable file (JSON or YAML)
        #[arg(short, long)]
        vars: Option<std::path::PathBuf>,
        /// Inline variable (key=value)
        #[arg(short = 'D', long = "define")]
        define: Vec<String>,
    },
    /// Initialize a new template with example placeholders
    Init {
        /// Output template file
        output: std::path::PathBuf,
        /// Template type (report, invoice, data)
        #[arg(short, long, default_value = "report")]
        template_type: String,
    },
    /// List placeholders in a template
    List {
        /// Template xlsx file
        template: std::path::PathBuf,
    },
    /// Validate template placeholders
    Validate {
        /// Template xlsx file
        template: std::path::PathBuf,
        /// Variable file (JSON or YAML)
        #[arg(short, long)]
        vars: Option<std::path::PathBuf>,
        /// Generate JSON schema for required data
        #[arg(long)]
        schema: bool,
    },
    /// Create template from existing file
    Create {
        /// Source xlsx file
        source: std::path::PathBuf,
        /// Output template file
        output: std::path::PathBuf,
        /// Cell to convert to placeholder (cell=name)
        #[arg(short, long)]
        placeholder: Vec<String>,
    },
    /// Preview template rendering without modifying files
    Preview {
        /// Template xlsx file
        template: std::path::PathBuf,
        /// Variable file (JSON or YAML)
        #[arg(short, long)]
        vars: Option<std::path::PathBuf>,
        /// Inline variable (key=value)
        #[arg(short = 'D', long = "define")]
        define: Vec<String>,
    },
}

/// Run template operations.
pub fn run(args: &TemplateArgs, global: &GlobalOptions) -> Result<()> {
    match &args.command {
        TemplateCommand::Apply {
            template,
            output,
            vars,
            define,
        } => apply(template, output, vars.as_deref(), define, global),
        TemplateCommand::Init {
            output,
            template_type,
        } => init(output, template_type, global),
        TemplateCommand::List { template } => list(template, global),
        TemplateCommand::Validate { template, vars, schema } => {
            validate(template, vars.as_deref(), *schema, global)
        }
        TemplateCommand::Create {
            source,
            output,
            placeholder,
        } => create(source, output, placeholder, global),
        TemplateCommand::Preview {
            template,
            vars,
            define,
        } => preview(template, vars.as_deref(), define, global),
    }
}

fn apply(
    template: &std::path::Path,
    output: &std::path::Path,
    vars_file: Option<&std::path::Path>,
    defines: &[String],
    global: &GlobalOptions,
) -> Result<()> {
    if global.dry_run {
        println!(
            "Would apply template {} to {}",
            template.display(),
            output.display()
        );
        return Ok(());
    }

    // Load variables
    let mut vars: HashMap<String, String> = HashMap::new();

    // Load from file
    if let Some(path) = vars_file {
        let content = std::fs::read_to_string(path)?;
        if path.extension().map(|e| e == "json").unwrap_or(false) {
            let json: serde_json::Value = serde_json::from_str(&content)?;
            if let serde_json::Value::Object(obj) = json {
                for (k, v) in obj {
                    vars.insert(k, value_to_string(&v));
                }
            }
        } else {
            let yaml: serde_yaml::Value = serde_yaml::from_str(&content)?;
            if let serde_yaml::Value::Mapping(map) = yaml {
                for (k, v) in map {
                    if let serde_yaml::Value::String(key) = k {
                        vars.insert(key, yaml_value_to_string(&v));
                    }
                }
            }
        }
    }

    // Add inline defines (override file vars)
    for define in defines {
        if let Some((key, value)) = define.split_once('=') {
            vars.insert(key.to_string(), value.to_string());
        }
    }

    // Open template and process
    let mut workbook = Workbook::open(template)?;

    let sheet_names: Vec<String> = workbook.sheet_names().iter().map(|s| s.to_string()).collect();
    for sheet_name in sheet_names {
        // Collect cell data we need before releasing the borrow
        let cells_to_update: Vec<(CellRef, String)> = if let Some(sheet) = workbook.get_sheet(&sheet_name) {
            sheet.cells()
                .filter_map(|cell| {
                    if let CellValue::String(s) = &cell.value {
                        let new_value = replace_placeholders(s, &vars);
                        if new_value != *s {
                            return Some((cell.reference.clone(), new_value));
                        }
                    }
                    None
                })
                .collect()
        } else {
            continue;
        };

        // Now apply the updates
        for (cell_ref, new_value) in cells_to_update {
            workbook.set_cell(
                &sheet_name,
                cell_ref,
                CellValue::String(new_value),
            )?;
        }
    }

    workbook.save_as(output)?;

    if !global.quiet {
        println!(
            "Applied template to {}",
            output.display().to_string().green()
        );
    }

    Ok(())
}

fn init(
    output: &std::path::Path,
    template_type: &str,
    global: &GlobalOptions,
) -> Result<()> {
    if global.dry_run {
        println!("Would create {} template at {}", template_type, output.display());
        return Ok(());
    }

    // Check if file exists
    if output.exists() {
        anyhow::bail!("File already exists: {}", output.display());
    }

    let mut workbook = Workbook::new();

    match template_type {
        "report" => {
            // Create a report template
            let sheet = workbook.get_sheet_mut("Sheet1").unwrap();
            sheet.set_cell(CellRef::parse("A1")?, CellValue::String("{{title}}".to_string()));
            sheet.set_cell(CellRef::parse("A2")?, CellValue::String("Date: {{date}}".to_string()));
            sheet.set_cell(CellRef::parse("A3")?, CellValue::String("Author: {{author}}".to_string()));
            sheet.set_cell(CellRef::parse("A5")?, CellValue::String("Summary".to_string()));
            sheet.set_cell(CellRef::parse("A6")?, CellValue::String("{{summary}}".to_string()));
            sheet.set_cell(CellRef::parse("A8")?, CellValue::String("Item".to_string()));
            sheet.set_cell(CellRef::parse("B8")?, CellValue::String("Value".to_string()));
            sheet.set_cell(CellRef::parse("A9")?, CellValue::String("{{item1_name}}".to_string()));
            sheet.set_cell(CellRef::parse("B9")?, CellValue::String("{{item1_value}}".to_string()));
            sheet.set_cell(CellRef::parse("A10")?, CellValue::String("{{item2_name}}".to_string()));
            sheet.set_cell(CellRef::parse("B10")?, CellValue::String("{{item2_value}}".to_string()));
        }
        "invoice" => {
            // Create an invoice template
            let sheet = workbook.get_sheet_mut("Sheet1").unwrap();
            sheet.set_cell(CellRef::parse("A1")?, CellValue::String("INVOICE".to_string()));
            sheet.set_cell(CellRef::parse("A3")?, CellValue::String("Invoice #: {{invoice_number}}".to_string()));
            sheet.set_cell(CellRef::parse("A4")?, CellValue::String("Date: {{invoice_date}}".to_string()));
            sheet.set_cell(CellRef::parse("A6")?, CellValue::String("Bill To:".to_string()));
            sheet.set_cell(CellRef::parse("A7")?, CellValue::String("{{customer_name}}".to_string()));
            sheet.set_cell(CellRef::parse("A8")?, CellValue::String("{{customer_address}}".to_string()));
            sheet.set_cell(CellRef::parse("A10")?, CellValue::String("Description".to_string()));
            sheet.set_cell(CellRef::parse("B10")?, CellValue::String("Qty".to_string()));
            sheet.set_cell(CellRef::parse("C10")?, CellValue::String("Price".to_string()));
            sheet.set_cell(CellRef::parse("D10")?, CellValue::String("Total".to_string()));
            sheet.set_cell(CellRef::parse("A11")?, CellValue::String("{{line1_desc}}".to_string()));
            sheet.set_cell(CellRef::parse("B11")?, CellValue::String("{{line1_qty}}".to_string()));
            sheet.set_cell(CellRef::parse("C11")?, CellValue::String("{{line1_price}}".to_string()));
            sheet.set_cell(CellRef::parse("D11")?, CellValue::String("{{line1_total}}".to_string()));
            sheet.set_cell(CellRef::parse("C14")?, CellValue::String("Subtotal:".to_string()));
            sheet.set_cell(CellRef::parse("D14")?, CellValue::String("{{subtotal}}".to_string()));
            sheet.set_cell(CellRef::parse("C15")?, CellValue::String("Tax:".to_string()));
            sheet.set_cell(CellRef::parse("D15")?, CellValue::String("{{tax}}".to_string()));
            sheet.set_cell(CellRef::parse("C16")?, CellValue::String("Total:".to_string()));
            sheet.set_cell(CellRef::parse("D16")?, CellValue::String("{{total}}".to_string()));
        }
        "data" => {
            // Create a simple data template
            let sheet = workbook.get_sheet_mut("Sheet1").unwrap();
            sheet.set_cell(CellRef::parse("A1")?, CellValue::String("{{header1}}".to_string()));
            sheet.set_cell(CellRef::parse("B1")?, CellValue::String("{{header2}}".to_string()));
            sheet.set_cell(CellRef::parse("C1")?, CellValue::String("{{header3}}".to_string()));
            sheet.set_cell(CellRef::parse("A2")?, CellValue::String("{{row1_col1}}".to_string()));
            sheet.set_cell(CellRef::parse("B2")?, CellValue::String("{{row1_col2}}".to_string()));
            sheet.set_cell(CellRef::parse("C2")?, CellValue::String("{{row1_col3}}".to_string()));
            sheet.set_cell(CellRef::parse("A3")?, CellValue::String("{{row2_col1}}".to_string()));
            sheet.set_cell(CellRef::parse("B3")?, CellValue::String("{{row2_col2}}".to_string()));
            sheet.set_cell(CellRef::parse("C3")?, CellValue::String("{{row2_col3}}".to_string()));
        }
        _ => {
            anyhow::bail!("Unknown template type: {}. Use: report, invoice, or data", template_type);
        }
    }

    workbook.save_as(output)?;

    if !global.quiet {
        if global.format == OutputFormat::Json {
            let json = serde_json::json!({
                "created": output.display().to_string(),
                "type": template_type,
            });
            println!("{}", serde_json::to_string_pretty(&json)?);
        } else {
            println!(
                "{} Created {} template: {}",
                "✓".green(),
                template_type.cyan(),
                output.display().to_string().yellow()
            );
            println!("\nUse {} to see placeholders", "xlex template list <file>".dimmed());
            println!("Use {} to apply variables", "xlex template apply <template> <output> -D key=value".dimmed());
        }
    }

    Ok(())
}

fn list(template: &std::path::Path, global: &GlobalOptions) -> Result<()> {
    let workbook = Workbook::open(template)?;

    let mut placeholders: Vec<(String, String, String)> = Vec::new(); // (sheet, cell, placeholder)

    for sheet_name in workbook.sheet_names() {
        if let Some(sheet) = workbook.get_sheet(sheet_name) {
            for cell in sheet.cells() {
                if let CellValue::String(s) = &cell.value {
                    for placeholder in find_placeholders(s) {
                        placeholders.push((
                            sheet_name.to_string(),
                            cell.reference.to_a1(),
                            placeholder,
                        ));
                    }
                }
            }
        }
    }

    // Get unique placeholder names
    let mut unique: Vec<String> = placeholders
        .iter()
        .map(|(_, _, p)| p.clone())
        .collect();
    unique.sort();
    unique.dedup();

    if global.format == OutputFormat::Json {
        let json = serde_json::json!({
            "placeholders": unique,
            "locations": placeholders.iter().map(|(s, c, p)| {
                serde_json::json!({
                    "sheet": s,
                    "cell": c,
                    "placeholder": p,
                })
            }).collect::<Vec<_>>(),
        });
        println!("{}", serde_json::to_string_pretty(&json)?);
    } else {
        println!("{}: {}", "Placeholders".bold(), unique.len());
        for name in &unique {
            println!("  {}", format!("{{{{{}}}}}", name).cyan());
        }

        if global.verbose {
            println!("\n{}:", "Locations".bold());
            for (sheet, cell, placeholder) in &placeholders {
                println!(
                    "  {}!{}: {}",
                    sheet,
                    cell.cyan(),
                    format!("{{{{{}}}}}", placeholder)
                );
            }
        }
    }

    Ok(())
}

fn validate(
    template: &std::path::Path,
    vars_file: Option<&std::path::Path>,
    generate_schema: bool,
    global: &GlobalOptions,
) -> Result<()> {
    let workbook = Workbook::open(template)?;

    // Find all placeholders
    let mut placeholders: Vec<String> = Vec::new();
    for sheet_name in workbook.sheet_names() {
        if let Some(sheet) = workbook.get_sheet(sheet_name) {
            for cell in sheet.cells() {
                if let CellValue::String(s) = &cell.value {
                    for placeholder in find_placeholders(s) {
                        if !placeholders.contains(&placeholder) {
                            placeholders.push(placeholder);
                        }
                    }
                }
            }
        }
    }

    // If --schema flag, generate JSON schema
    if generate_schema {
        let mut properties = serde_json::Map::new();
        for p in &placeholders {
            properties.insert(
                p.clone(),
                serde_json::json!({
                    "type": "string",
                    "description": format!("Value for placeholder {{{{{}}}}}", p)
                }),
            );
        }
        
        let schema = serde_json::json!({
            "$schema": "http://json-schema.org/draft-07/schema#",
            "type": "object",
            "properties": properties,
            "required": placeholders,
        });
        
        println!("{}", serde_json::to_string_pretty(&schema)?);
        return Ok(());
    }

    // Load available variables
    let mut vars: Vec<String> = Vec::new();
    if let Some(path) = vars_file {
        let content = std::fs::read_to_string(path)?;
        if path.extension().map(|e| e == "json").unwrap_or(false) {
            let json: serde_json::Value = serde_json::from_str(&content)?;
            if let serde_json::Value::Object(obj) = json {
                vars = obj.keys().cloned().collect();
            }
        } else {
            let yaml: serde_yaml::Value = serde_yaml::from_str(&content)?;
            if let serde_yaml::Value::Mapping(map) = yaml {
                for k in map.keys() {
                    if let serde_yaml::Value::String(key) = k {
                        vars.push(key.clone());
                    }
                }
            }
        }
    }

    let missing: Vec<_> = placeholders
        .iter()
        .filter(|p| !vars.contains(p))
        .collect();

    let unused: Vec<_> = vars.iter().filter(|v| !placeholders.contains(v)).collect();

    if global.format == OutputFormat::Json {
        let json = serde_json::json!({
            "valid": missing.is_empty(),
            "placeholders": placeholders,
            "variables": vars,
            "missing": missing,
            "unused": unused,
        });
        println!("{}", serde_json::to_string_pretty(&json)?);
    } else {
        if missing.is_empty() {
            println!("{}: All placeholders have values", "✓".green());
        } else {
            println!("{}: Missing values for:", "✗".red());
            for m in &missing {
                println!("  - {}", format!("{{{{{}}}}}", m).cyan());
            }
        }

        if !unused.is_empty() && global.verbose {
            println!("\n{}: Unused variables:", "⚠".yellow());
            for u in &unused {
                println!("  - {}", u);
            }
        }
    }

    Ok(())
}

fn preview(
    template: &std::path::Path,
    vars_file: Option<&std::path::Path>,
    defines: &[String],
    global: &GlobalOptions,
) -> Result<()> {
    // Load variables
    let mut vars: HashMap<String, String> = HashMap::new();

    // Load from file
    if let Some(path) = vars_file {
        let content = std::fs::read_to_string(path)?;
        if path.extension().map(|e| e == "json").unwrap_or(false) {
            let json: serde_json::Value = serde_json::from_str(&content)?;
            if let serde_json::Value::Object(obj) = json {
                for (k, v) in obj {
                    vars.insert(k, value_to_string(&v));
                }
            }
        } else {
            let yaml: serde_yaml::Value = serde_yaml::from_str(&content)?;
            if let serde_yaml::Value::Mapping(map) = yaml {
                for (k, v) in map {
                    if let serde_yaml::Value::String(key) = k {
                        vars.insert(key, yaml_value_to_string(&v));
                    }
                }
            }
        }
    }

    // Add inline defines
    for define in defines {
        if let Some((key, value)) = define.split_once('=') {
            vars.insert(key.to_string(), value.to_string());
        }
    }

    let workbook = Workbook::open(template)?;

    // Collect all replacements
    let mut replacements: Vec<serde_json::Value> = Vec::new();

    for sheet_name in workbook.sheet_names() {
        if let Some(sheet) = workbook.get_sheet(sheet_name) {
            for cell in sheet.cells() {
                if let CellValue::String(s) = &cell.value {
                    let new_value = replace_placeholders(s, &vars);
                    if new_value != *s {
                        replacements.push(serde_json::json!({
                            "sheet": sheet_name,
                            "cell": cell.reference.to_a1(),
                            "original": s,
                            "result": new_value,
                        }));
                    }
                }
            }
        }
    }

    if global.format == OutputFormat::Json {
        let json = serde_json::json!({
            "template": template.display().to_string(),
            "variables": vars,
            "replacements": replacements,
        });
        println!("{}", serde_json::to_string_pretty(&json)?);
    } else {
        println!("{}: {}\n", "Template".bold(), template.display());
        
        if vars.is_empty() {
            println!("{}", "No variables provided".yellow());
        } else {
            println!("{} ({})", "Variables".bold(), vars.len());
            for (k, v) in &vars {
                println!("  {} = {}", k.cyan(), v.green());
            }
            println!();
        }

        if replacements.is_empty() {
            println!("{}", "No replacements would be made".dimmed());
        } else {
            println!("{} ({}):", "Replacements".bold(), replacements.len());
            for r in &replacements {
                println!(
                    "  [{}] {}: {} → {}",
                    r["sheet"].as_str().unwrap(),
                    r["cell"].as_str().unwrap().cyan(),
                    r["original"].as_str().unwrap().dimmed(),
                    r["result"].as_str().unwrap().green()
                );
            }
        }
    }

    Ok(())
}

fn create(
    source: &std::path::Path,
    output: &std::path::Path,
    placeholders: &[String],
    global: &GlobalOptions,
) -> Result<()> {
    if global.dry_run {
        println!(
            "Would create template {} from {}",
            output.display(),
            source.display()
        );
        return Ok(());
    }

    let mut workbook = Workbook::open(source)?;

    for placeholder in placeholders {
        if let Some((cell_str, name)) = placeholder.split_once('=') {
            // Parse cell reference (may include sheet: Sheet1!A1)
            let (sheet_name, cell_ref) = if cell_str.contains('!') {
                let parts: Vec<_> = cell_str.splitn(2, '!').collect();
                (parts[0].to_string(), CellRef::parse(parts[1])?)
            } else {
                let first_sheet = workbook
                    .sheet_names()
                    .first()
                    .copied()
                    .unwrap_or("Sheet1");
                (first_sheet.to_string(), CellRef::parse(cell_str)?)
            };

            workbook.set_cell(
                &sheet_name,
                cell_ref,
                CellValue::String(format!("{{{{{}}}}}", name)),
            )?;
        }
    }

    workbook.save_as(output)?;

    if !global.quiet {
        println!(
            "Created template {} with {} placeholders",
            output.display().to_string().green(),
            placeholders.len()
        );
    }

    Ok(())
}

fn find_placeholders(s: &str) -> Vec<String> {
    let mut result = Vec::new();
    let mut chars = s.chars().peekable();
    let mut current = String::new();
    let mut in_placeholder = false;

    while let Some(c) = chars.next() {
        if c == '{' && chars.peek() == Some(&'{') {
            chars.next(); // consume second {
            in_placeholder = true;
            current.clear();
        } else if c == '}' && in_placeholder && chars.peek() == Some(&'}') {
            chars.next(); // consume second }
            in_placeholder = false;
            if !current.is_empty() {
                result.push(current.clone());
            }
        } else if in_placeholder {
            current.push(c);
        }
    }

    result
}

fn replace_placeholders(s: &str, vars: &HashMap<String, String>) -> String {
    let mut result = s.to_string();
    for (key, value) in vars {
        let placeholder = format!("{{{{{}}}}}", key);
        result = result.replace(&placeholder, value);
    }
    result
}

fn value_to_string(v: &serde_json::Value) -> String {
    match v {
        serde_json::Value::String(s) => s.clone(),
        serde_json::Value::Number(n) => n.to_string(),
        serde_json::Value::Bool(b) => b.to_string(),
        serde_json::Value::Null => String::new(),
        _ => v.to_string(),
    }
}

fn yaml_value_to_string(v: &serde_yaml::Value) -> String {
    match v {
        serde_yaml::Value::String(s) => s.clone(),
        serde_yaml::Value::Number(n) => {
            if let Some(i) = n.as_i64() {
                i.to_string()
            } else if let Some(f) = n.as_f64() {
                f.to_string()
            } else {
                String::new()
            }
        }
        serde_yaml::Value::Bool(b) => b.to_string(),
        serde_yaml::Value::Null => String::new(),
        _ => format!("{:?}", v),
    }
}
