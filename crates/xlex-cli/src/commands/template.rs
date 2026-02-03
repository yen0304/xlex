//! Template operations.

use anyhow::Result;
use clap::{Parser, Subcommand};
use colored::Colorize;

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
        #[arg(long)]
        vars: Option<std::path::PathBuf>,
        /// Inline variable (key=value)
        #[arg(short = 'D', long = "define")]
        define: Vec<String>,
        /// Generate one file per record (for array data)
        #[arg(long)]
        per_record: bool,
        /// Output pattern for per-record mode (e.g., "output_{index}.xlsx")
        #[arg(long)]
        output_pattern: Option<String>,
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
        #[arg(long)]
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
        #[arg(long)]
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
            per_record,
            output_pattern,
        } => apply(
            template,
            output,
            vars.as_deref(),
            define,
            *per_record,
            output_pattern.as_deref(),
            global,
        ),
        TemplateCommand::Init {
            output,
            template_type,
        } => init(output, template_type, global),
        TemplateCommand::List { template } => list(template, global),
        TemplateCommand::Validate {
            template,
            vars,
            schema,
        } => validate(template, vars.as_deref(), *schema, global),
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
    per_record: bool,
    output_pattern: Option<&str>,
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
    let vars = load_template_vars(vars_file, defines)?;

    // Handle per-record mode for batch processing
    if per_record {
        return apply_per_record(template, output, output_pattern, &vars, global);
    }

    // Single file processing with advanced template features
    apply_single(template, output, &vars, global)
}

/// Apply template to generate a single output file.
fn apply_single(
    template: &std::path::Path,
    output: &std::path::Path,
    vars: &TemplateVars,
    global: &GlobalOptions,
) -> Result<()> {
    let mut workbook = Workbook::open(template)?;

    let sheet_names: Vec<String> = workbook
        .sheet_names()
        .iter()
        .map(|s| s.to_string())
        .collect();

    for sheet_name in sheet_names {
        // First, handle row-repeat markers
        process_row_repeats(&mut workbook, &sheet_name, vars)?;

        // Then, process all cells for placeholders
        let cells_to_update: Vec<(CellRef, String)> =
            if let Some(sheet) = workbook.get_sheet(&sheet_name) {
                sheet
                    .cells()
                    .filter_map(|cell| {
                        if let CellValue::String(s) = &cell.value {
                            let new_value = process_template_string(s, vars);
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

        // Apply the updates
        for (cell_ref, new_value) in cells_to_update {
            workbook.set_cell(&sheet_name, cell_ref, CellValue::String(new_value))?;
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

/// Apply template per-record for batch processing.
fn apply_per_record(
    template: &std::path::Path,
    output: &std::path::Path,
    output_pattern: Option<&str>,
    vars: &TemplateVars,
    global: &GlobalOptions,
) -> Result<()> {
    // Check if vars contains an array (records)
    let records = vars
        .get_array("records")
        .or_else(|| vars.get_array("items"))
        .or_else(|| vars.get_array("data"));

    let records = match records {
        Some(r) => r,
        None => {
            anyhow::bail!("Per-record mode requires an array named 'records', 'items', or 'data' in the variables");
        }
    };

    let pattern = output_pattern.unwrap_or("{name}_{index}.xlsx");
    let output_dir = output.parent().unwrap_or(std::path::Path::new("."));
    let base_name = output
        .file_stem()
        .and_then(|s| s.to_str())
        .unwrap_or("output");

    let mut generated_files = Vec::new();

    for (index, record) in records.iter().enumerate() {
        // Create vars for this record
        let mut record_vars = vars.clone();
        record_vars.merge_object(record);
        record_vars.set("index", &(index + 1).to_string());
        record_vars.set("index0", &index.to_string());

        // Generate output filename
        let output_name = pattern
            .replace("{index}", &(index + 1).to_string())
            .replace("{index0}", &index.to_string())
            .replace("{name}", base_name);

        let output_path = output_dir.join(&output_name);

        apply_single(
            template,
            &output_path,
            &record_vars,
            &GlobalOptions {
                quiet: true,
                ..global.clone()
            },
        )?;

        generated_files.push(output_path);
    }

    if !global.quiet {
        if global.format == OutputFormat::Json {
            let json = serde_json::json!({
                "generated": generated_files.iter().map(|p| p.display().to_string()).collect::<Vec<_>>(),
                "count": generated_files.len(),
            });
            println!("{}", serde_json::to_string_pretty(&json)?);
        } else {
            println!(
                "{} Generated {} files from template",
                "✓".green(),
                generated_files.len().to_string().cyan()
            );
            for path in &generated_files {
                println!("  - {}", path.display().to_string().yellow());
            }
        }
    }

    Ok(())
}

fn init(output: &std::path::Path, template_type: &str, global: &GlobalOptions) -> Result<()> {
    if global.dry_run {
        println!(
            "Would create {} template at {}",
            template_type,
            output.display()
        );
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
            sheet.set_cell(
                CellRef::parse("A1")?,
                CellValue::String("{{title}}".to_string()),
            );
            sheet.set_cell(
                CellRef::parse("A2")?,
                CellValue::String("Date: {{date}}".to_string()),
            );
            sheet.set_cell(
                CellRef::parse("A3")?,
                CellValue::String("Author: {{author}}".to_string()),
            );
            sheet.set_cell(
                CellRef::parse("A5")?,
                CellValue::String("Summary".to_string()),
            );
            sheet.set_cell(
                CellRef::parse("A6")?,
                CellValue::String("{{summary}}".to_string()),
            );
            sheet.set_cell(CellRef::parse("A8")?, CellValue::String("Item".to_string()));
            sheet.set_cell(
                CellRef::parse("B8")?,
                CellValue::String("Value".to_string()),
            );
            sheet.set_cell(
                CellRef::parse("A9")?,
                CellValue::String("{{item1_name}}".to_string()),
            );
            sheet.set_cell(
                CellRef::parse("B9")?,
                CellValue::String("{{item1_value}}".to_string()),
            );
            sheet.set_cell(
                CellRef::parse("A10")?,
                CellValue::String("{{item2_name}}".to_string()),
            );
            sheet.set_cell(
                CellRef::parse("B10")?,
                CellValue::String("{{item2_value}}".to_string()),
            );
        }
        "invoice" => {
            // Create an invoice template
            let sheet = workbook.get_sheet_mut("Sheet1").unwrap();
            sheet.set_cell(
                CellRef::parse("A1")?,
                CellValue::String("INVOICE".to_string()),
            );
            sheet.set_cell(
                CellRef::parse("A3")?,
                CellValue::String("Invoice #: {{invoice_number}}".to_string()),
            );
            sheet.set_cell(
                CellRef::parse("A4")?,
                CellValue::String("Date: {{invoice_date}}".to_string()),
            );
            sheet.set_cell(
                CellRef::parse("A6")?,
                CellValue::String("Bill To:".to_string()),
            );
            sheet.set_cell(
                CellRef::parse("A7")?,
                CellValue::String("{{customer_name}}".to_string()),
            );
            sheet.set_cell(
                CellRef::parse("A8")?,
                CellValue::String("{{customer_address}}".to_string()),
            );
            sheet.set_cell(
                CellRef::parse("A10")?,
                CellValue::String("Description".to_string()),
            );
            sheet.set_cell(CellRef::parse("B10")?, CellValue::String("Qty".to_string()));
            sheet.set_cell(
                CellRef::parse("C10")?,
                CellValue::String("Price".to_string()),
            );
            sheet.set_cell(
                CellRef::parse("D10")?,
                CellValue::String("Total".to_string()),
            );
            sheet.set_cell(
                CellRef::parse("A11")?,
                CellValue::String("{{line1_desc}}".to_string()),
            );
            sheet.set_cell(
                CellRef::parse("B11")?,
                CellValue::String("{{line1_qty}}".to_string()),
            );
            sheet.set_cell(
                CellRef::parse("C11")?,
                CellValue::String("{{line1_price}}".to_string()),
            );
            sheet.set_cell(
                CellRef::parse("D11")?,
                CellValue::String("{{line1_total}}".to_string()),
            );
            sheet.set_cell(
                CellRef::parse("C14")?,
                CellValue::String("Subtotal:".to_string()),
            );
            sheet.set_cell(
                CellRef::parse("D14")?,
                CellValue::String("{{subtotal}}".to_string()),
            );
            sheet.set_cell(
                CellRef::parse("C15")?,
                CellValue::String("Tax:".to_string()),
            );
            sheet.set_cell(
                CellRef::parse("D15")?,
                CellValue::String("{{tax}}".to_string()),
            );
            sheet.set_cell(
                CellRef::parse("C16")?,
                CellValue::String("Total:".to_string()),
            );
            sheet.set_cell(
                CellRef::parse("D16")?,
                CellValue::String("{{total}}".to_string()),
            );
        }
        "data" => {
            // Create a simple data template
            let sheet = workbook.get_sheet_mut("Sheet1").unwrap();
            sheet.set_cell(
                CellRef::parse("A1")?,
                CellValue::String("{{header1}}".to_string()),
            );
            sheet.set_cell(
                CellRef::parse("B1")?,
                CellValue::String("{{header2}}".to_string()),
            );
            sheet.set_cell(
                CellRef::parse("C1")?,
                CellValue::String("{{header3}}".to_string()),
            );
            sheet.set_cell(
                CellRef::parse("A2")?,
                CellValue::String("{{row1_col1}}".to_string()),
            );
            sheet.set_cell(
                CellRef::parse("B2")?,
                CellValue::String("{{row1_col2}}".to_string()),
            );
            sheet.set_cell(
                CellRef::parse("C2")?,
                CellValue::String("{{row1_col3}}".to_string()),
            );
            sheet.set_cell(
                CellRef::parse("A3")?,
                CellValue::String("{{row2_col1}}".to_string()),
            );
            sheet.set_cell(
                CellRef::parse("B3")?,
                CellValue::String("{{row2_col2}}".to_string()),
            );
            sheet.set_cell(
                CellRef::parse("C3")?,
                CellValue::String("{{row2_col3}}".to_string()),
            );
        }
        _ => {
            anyhow::bail!(
                "Unknown template type: {}. Use: report, invoice, or data",
                template_type
            );
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
            println!(
                "\nUse {} to see placeholders",
                "xlex template list <file>".dimmed()
            );
            println!(
                "Use {} to apply variables",
                "xlex template apply <template> <output> -D key=value".dimmed()
            );
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
    let mut unique: Vec<String> = placeholders.iter().map(|(_, _, p)| p.clone()).collect();
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

    let missing: Vec<_> = placeholders.iter().filter(|p| !vars.contains(p)).collect();

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
    // Load variables using the new template system
    let vars = load_template_vars(vars_file, defines)?;

    let workbook = Workbook::open(template)?;

    // Collect all replacements using the advanced template processor
    let mut replacements: Vec<serde_json::Value> = Vec::new();

    for sheet_name in workbook.sheet_names() {
        if let Some(sheet) = workbook.get_sheet(sheet_name) {
            for cell in sheet.cells() {
                if let CellValue::String(s) = &cell.value {
                    let new_value = process_template_string(s, &vars);
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
            "replacements": replacements,
        });
        println!("{}", serde_json::to_string_pretty(&json)?);
    } else {
        println!("{}: {}\n", "Template".bold(), template.display());

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
                let first_sheet = workbook.sheet_names().first().copied().unwrap_or("Sheet1");
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

fn value_to_string(v: &serde_json::Value) -> String {
    match v {
        serde_json::Value::String(s) => s.clone(),
        serde_json::Value::Number(n) => n.to_string(),
        serde_json::Value::Bool(b) => b.to_string(),
        serde_json::Value::Null => String::new(),
        _ => v.to_string(),
    }
}

// =============================================================================
// Advanced Template Engine
// =============================================================================

/// Template variables container supporting nested access and arrays.
#[derive(Clone, Debug, Default)]
pub struct TemplateVars {
    data: serde_json::Value,
}

impl TemplateVars {
    pub fn new() -> Self {
        Self {
            data: serde_json::Value::Object(serde_json::Map::new()),
        }
    }

    pub fn from_json(value: serde_json::Value) -> Self {
        Self { data: value }
    }

    pub fn set(&mut self, key: &str, value: &str) {
        if let serde_json::Value::Object(ref mut map) = self.data {
            map.insert(
                key.to_string(),
                serde_json::Value::String(value.to_string()),
            );
        }
    }

    pub fn get(&self, key: &str) -> Option<String> {
        // Support dot notation for nested access (e.g., "user.name")
        let parts: Vec<&str> = key.split('.').collect();
        let mut current = &self.data;

        for part in &parts {
            match current {
                serde_json::Value::Object(map) => {
                    current = map.get(*part)?;
                }
                serde_json::Value::Array(arr) => {
                    let idx: usize = part.parse().ok()?;
                    current = arr.get(idx)?;
                }
                _ => return None,
            }
        }

        Some(value_to_string(current))
    }

    pub fn get_array(&self, key: &str) -> Option<Vec<serde_json::Value>> {
        let parts: Vec<&str> = key.split('.').collect();
        let mut current = &self.data;

        for part in &parts {
            match current {
                serde_json::Value::Object(map) => {
                    current = map.get(*part)?;
                }
                _ => return None,
            }
        }

        if let serde_json::Value::Array(arr) = current {
            Some(arr.clone())
        } else {
            None
        }
    }

    pub fn get_bool(&self, key: &str) -> Option<bool> {
        let parts: Vec<&str> = key.split('.').collect();
        let mut current = &self.data;

        for part in &parts {
            if let serde_json::Value::Object(map) = current {
                current = map.get(*part)?;
            } else {
                return None;
            }
        }

        match current {
            serde_json::Value::Bool(b) => Some(*b),
            serde_json::Value::String(s) => match s.to_lowercase().as_str() {
                "true" | "yes" | "1" => Some(true),
                "false" | "no" | "0" => Some(false),
                _ => Some(!s.is_empty()),
            },
            serde_json::Value::Number(n) => Some(n.as_f64().map(|f| f != 0.0).unwrap_or(false)),
            serde_json::Value::Null => Some(false),
            serde_json::Value::Array(arr) => Some(!arr.is_empty()),
            serde_json::Value::Object(obj) => Some(!obj.is_empty()),
        }
    }

    pub fn merge_object(&mut self, obj: &serde_json::Value) {
        if let (serde_json::Value::Object(ref mut target), serde_json::Value::Object(source)) =
            (&mut self.data, obj)
        {
            for (k, v) in source {
                target.insert(k.clone(), v.clone());
            }
        }
    }
}

/// Load template variables from file and command-line defines.
fn load_template_vars(
    vars_file: Option<&std::path::Path>,
    defines: &[String],
) -> Result<TemplateVars> {
    let mut vars = TemplateVars::new();

    // Load from file
    if let Some(path) = vars_file {
        let content = std::fs::read_to_string(path)?;
        let json_value = if path.extension().map(|e| e == "json").unwrap_or(false) {
            serde_json::from_str(&content)?
        } else {
            // Parse YAML and convert to JSON
            let yaml: serde_yaml::Value = serde_yaml::from_str(&content)?;
            yaml_to_json(&yaml)
        };
        vars = TemplateVars::from_json(json_value);
    }

    // Add inline defines (override file vars)
    for define in defines {
        if let Some((key, value)) = define.split_once('=') {
            vars.set(key, value);
        }
    }

    Ok(vars)
}

/// Convert YAML value to JSON value.
fn yaml_to_json(yaml: &serde_yaml::Value) -> serde_json::Value {
    match yaml {
        serde_yaml::Value::Null => serde_json::Value::Null,
        serde_yaml::Value::Bool(b) => serde_json::Value::Bool(*b),
        serde_yaml::Value::Number(n) => {
            if let Some(i) = n.as_i64() {
                serde_json::Value::Number(i.into())
            } else if let Some(f) = n.as_f64() {
                serde_json::Number::from_f64(f)
                    .map(serde_json::Value::Number)
                    .unwrap_or(serde_json::Value::Null)
            } else {
                serde_json::Value::Null
            }
        }
        serde_yaml::Value::String(s) => serde_json::Value::String(s.clone()),
        serde_yaml::Value::Sequence(arr) => {
            serde_json::Value::Array(arr.iter().map(yaml_to_json).collect())
        }
        serde_yaml::Value::Mapping(map) => {
            let mut obj = serde_json::Map::new();
            for (k, v) in map {
                if let serde_yaml::Value::String(key) = k {
                    obj.insert(key.clone(), yaml_to_json(v));
                }
            }
            serde_json::Value::Object(obj)
        }
        serde_yaml::Value::Tagged(tagged) => yaml_to_json(&tagged.value),
    }
}

/// Process a template string with all advanced features.
fn process_template_string(s: &str, vars: &TemplateVars) -> String {
    let mut result = s.to_string();

    // Process conditionals first: {{#if condition}}...{{/if}}
    result = process_conditionals(&result, vars);

    // Process simple placeholders with filters: {{name|filter}}
    result = process_placeholders_with_filters(&result, vars);

    result
}

/// Process conditional blocks: {{#if condition}}content{{/if}} and {{#if condition}}content{{else}}other{{/if}}
fn process_conditionals(s: &str, vars: &TemplateVars) -> String {
    let mut result = s.to_string();

    // Pattern for {{#if condition}}...{{/if}} (with optional {{else}})
    let if_pattern = regex_lite::Regex::new(r"\{\{#if\s+([^}]+)\}\}([\s\S]*?)\{\{/if\}\}").unwrap();

    while let Some(captures) = if_pattern.captures(&result) {
        let full_match = captures.get(0).unwrap();
        let condition = captures.get(1).unwrap().as_str().trim();
        let content = captures.get(2).unwrap().as_str();

        // Check for else clause
        let (if_content, else_content) = if let Some(else_pos) = content.find("{{else}}") {
            (&content[..else_pos], &content[else_pos + 8..])
        } else {
            (content, "")
        };

        // Evaluate condition
        let is_true = evaluate_condition(condition, vars);

        let replacement = if is_true {
            if_content.to_string()
        } else {
            else_content.to_string()
        };

        result = result.replace(full_match.as_str(), &replacement);
    }

    // Also handle {{#unless condition}}...{{/unless}}
    let unless_pattern =
        regex_lite::Regex::new(r"\{\{#unless\s+([^}]+)\}\}([\s\S]*?)\{\{/unless\}\}").unwrap();

    while let Some(captures) = unless_pattern.captures(&result) {
        let full_match = captures.get(0).unwrap();
        let condition = captures.get(1).unwrap().as_str().trim();
        let content = captures.get(2).unwrap().as_str();

        let is_true = evaluate_condition(condition, vars);
        let replacement = if !is_true { content } else { "" };

        result = result.replace(full_match.as_str(), replacement);
    }

    result
}

/// Evaluate a condition expression.
fn evaluate_condition(condition: &str, vars: &TemplateVars) -> bool {
    // Handle comparison operators
    if condition.contains("==") {
        let parts: Vec<&str> = condition.split("==").collect();
        if parts.len() == 2 {
            let left = resolve_value(parts[0].trim(), vars);
            let right = resolve_value(parts[1].trim(), vars);
            return left == right;
        }
    }
    if condition.contains("!=") {
        let parts: Vec<&str> = condition.split("!=").collect();
        if parts.len() == 2 {
            let left = resolve_value(parts[0].trim(), vars);
            let right = resolve_value(parts[1].trim(), vars);
            return left != right;
        }
    }
    if condition.contains(">=") {
        let parts: Vec<&str> = condition.split(">=").collect();
        if parts.len() == 2 {
            let left = resolve_number(parts[0].trim(), vars);
            let right = resolve_number(parts[1].trim(), vars);
            if let (Some(l), Some(r)) = (left, right) {
                return l >= r;
            }
        }
    }
    if condition.contains("<=") {
        let parts: Vec<&str> = condition.split("<=").collect();
        if parts.len() == 2 {
            let left = resolve_number(parts[0].trim(), vars);
            let right = resolve_number(parts[1].trim(), vars);
            if let (Some(l), Some(r)) = (left, right) {
                return l <= r;
            }
        }
    }
    if condition.contains('>') && !condition.contains(">=") {
        let parts: Vec<&str> = condition.split('>').collect();
        if parts.len() == 2 {
            let left = resolve_number(parts[0].trim(), vars);
            let right = resolve_number(parts[1].trim(), vars);
            if let (Some(l), Some(r)) = (left, right) {
                return l > r;
            }
        }
    }
    if condition.contains('<') && !condition.contains("<=") {
        let parts: Vec<&str> = condition.split('<').collect();
        if parts.len() == 2 {
            let left = resolve_number(parts[0].trim(), vars);
            let right = resolve_number(parts[1].trim(), vars);
            if let (Some(l), Some(r)) = (left, right) {
                return l < r;
            }
        }
    }

    // Simple boolean check
    vars.get_bool(condition).unwrap_or(false)
}

fn resolve_value(expr: &str, vars: &TemplateVars) -> String {
    // Check if it's a quoted string
    if (expr.starts_with('"') && expr.ends_with('"'))
        || (expr.starts_with('\'') && expr.ends_with('\''))
    {
        return expr[1..expr.len() - 1].to_string();
    }
    // Otherwise resolve as variable
    vars.get(expr).unwrap_or_default()
}

fn resolve_number(expr: &str, vars: &TemplateVars) -> Option<f64> {
    let value = resolve_value(expr, vars);
    value.parse().ok().or_else(|| expr.parse().ok())
}

/// Process placeholders with optional filters: {{name}}, {{name|upper}}, {{amount|currency}}
fn process_placeholders_with_filters(s: &str, vars: &TemplateVars) -> String {
    let placeholder_pattern = regex_lite::Regex::new(r"\{\{([^#/][^}]*)\}\}").unwrap();

    let mut result = s.to_string();

    // Find all placeholders and process them
    for captures in placeholder_pattern.captures_iter(s) {
        let full_match = captures.get(0).unwrap().as_str();
        let inner = captures.get(1).unwrap().as_str().trim();

        // Parse variable name and filters
        let parts: Vec<&str> = inner.split('|').collect();
        let var_name = parts[0].trim();
        let filters: Vec<&str> = parts[1..].iter().map(|s| s.trim()).collect();

        // Get the value
        let mut value = vars.get(var_name).unwrap_or_else(|| full_match.to_string());

        // Apply filters
        for filter in filters {
            value = apply_filter(&value, filter);
        }

        result = result.replace(full_match, &value);
    }

    result
}

/// Apply a filter to a value.
fn apply_filter(value: &str, filter: &str) -> String {
    // Parse filter name and optional argument: filter or filter:arg
    let (filter_name, filter_arg) = if let Some(pos) = filter.find(':') {
        (&filter[..pos], Some(&filter[pos + 1..]))
    } else {
        (filter, None)
    };

    match filter_name {
        // String filters
        "upper" | "uppercase" => value.to_uppercase(),
        "lower" | "lowercase" => value.to_lowercase(),
        "capitalize" | "title" => {
            let mut chars = value.chars();
            match chars.next() {
                None => String::new(),
                Some(c) => c.to_uppercase().collect::<String>() + chars.as_str(),
            }
        }
        "trim" => value.trim().to_string(),
        "default" => {
            if value.is_empty() {
                filter_arg.unwrap_or("").to_string()
            } else {
                value.to_string()
            }
        }
        "truncate" => {
            let len: usize = filter_arg.and_then(|a| a.parse().ok()).unwrap_or(50);
            if value.chars().count() > len {
                format!("{}...", value.chars().take(len).collect::<String>())
            } else {
                value.to_string()
            }
        }
        "replace" => {
            // replace:old:new
            if let Some(arg) = filter_arg {
                let parts: Vec<&str> = arg.splitn(2, ':').collect();
                if parts.len() == 2 {
                    value.replace(parts[0], parts[1])
                } else {
                    value.to_string()
                }
            } else {
                value.to_string()
            }
        }

        // Number filters
        "currency" => {
            if let Ok(num) = value.parse::<f64>() {
                let symbol = filter_arg.unwrap_or("$");
                format!("{}{:.2}", symbol, num)
            } else {
                value.to_string()
            }
        }
        "number" | "format_number" => {
            if let Ok(num) = value.parse::<f64>() {
                let decimals: usize = filter_arg.and_then(|a| a.parse().ok()).unwrap_or(2);
                format!("{:.prec$}", num, prec = decimals)
            } else {
                value.to_string()
            }
        }
        "percent" => {
            if let Ok(num) = value.parse::<f64>() {
                format!("{:.1}%", num * 100.0)
            } else {
                value.to_string()
            }
        }
        "abs" => {
            if let Ok(num) = value.parse::<f64>() {
                num.abs().to_string()
            } else {
                value.to_string()
            }
        }
        "round" => {
            if let Ok(num) = value.parse::<f64>() {
                let decimals: i32 = filter_arg.and_then(|a| a.parse().ok()).unwrap_or(0);
                let factor = 10_f64.powi(decimals);
                ((num * factor).round() / factor).to_string()
            } else {
                value.to_string()
            }
        }

        // Date filters
        "date" => {
            // Try to parse and format date
            let format = filter_arg.unwrap_or("%Y-%m-%d");
            if let Ok(date) = chrono::NaiveDate::parse_from_str(value, "%Y-%m-%d") {
                date.format(format).to_string()
            } else if let Ok(datetime) = chrono::DateTime::parse_from_rfc3339(value) {
                datetime.format(format).to_string()
            } else {
                value.to_string()
            }
        }
        "now" => {
            let format = filter_arg.unwrap_or("%Y-%m-%d");
            chrono::Local::now().format(format).to_string()
        }

        // Unknown filter - return value unchanged
        _ => value.to_string(),
    }
}

/// Process row-repeat markers in a sheet.
/// Format: {{#row-repeat items}}...{{/row-repeat}}
fn process_row_repeats(
    workbook: &mut Workbook,
    sheet_name: &str,
    vars: &TemplateVars,
) -> Result<()> {
    // Find cells with row-repeat markers
    let repeat_markers: Vec<(u32, String, String)> =
        if let Some(sheet) = workbook.get_sheet(sheet_name) {
            sheet
                .cells()
                .filter_map(|cell| {
                    if let CellValue::String(s) = &cell.value {
                        // Look for {{#row-repeat array_name}}
                        if let Some(start) = s.find("{{#row-repeat") {
                            if let Some(end) = s[start..].find("}}") {
                                let marker = &s[start..start + end + 2];
                                // Extract array name
                                let inner = &marker[13..marker.len() - 2].trim();
                                return Some((cell.reference.row, inner.to_string(), s.clone()));
                            }
                        }
                    }
                    None
                })
                .collect()
        } else {
            return Ok(());
        };

    // Process each row-repeat marker
    for (row, array_name, _original) in repeat_markers.iter().rev() {
        // Get the array data
        if let Some(array) = vars.get_array(array_name) {
            if array.is_empty() {
                continue;
            }

            // Get the template row cells
            let template_cells: Vec<(u32, String)> =
                if let Some(sheet) = workbook.get_sheet(sheet_name) {
                    sheet
                        .cells()
                        .filter(|c| c.reference.row == *row)
                        .filter_map(|c| {
                            if let CellValue::String(s) = &c.value {
                                Some((c.reference.col, s.clone()))
                            } else {
                                None
                            }
                        })
                        .collect()
                } else {
                    continue;
                };

            // Generate rows for each array item
            for (idx, item) in array.iter().enumerate() {
                let target_row = *row + idx as u32;

                // Create vars for this item
                let mut item_vars = vars.clone();
                item_vars.merge_object(item);
                item_vars.set("_index", &(idx + 1).to_string());
                item_vars.set("_index0", &idx.to_string());
                item_vars.set("_first", &(idx == 0).to_string());
                item_vars.set("_last", &(idx == array.len() - 1).to_string());

                for (col, template) in &template_cells {
                    // Remove the row-repeat markers from template
                    let clean_template = template
                        .replace(&format!("{{{{#row-repeat {}}}}}", array_name), "")
                        .replace("{{/row-repeat}}", "");

                    let processed = process_template_string(&clean_template, &item_vars);

                    workbook.set_cell(
                        sheet_name,
                        CellRef::new(*col, target_row),
                        CellValue::String(processed),
                    )?;
                }
            }
        }
    }

    Ok(())
}

/// Find placeholders in a string (simple {{name}} format only).
fn find_placeholders(s: &str) -> Vec<String> {
    let mut result = Vec::new();
    let pattern = regex_lite::Regex::new(r"\{\{([^#/}][^}]*)\}\}").unwrap();

    for captures in pattern.captures_iter(s) {
        let inner = captures.get(1).unwrap().as_str().trim();
        // Extract just the variable name (without filters)
        let var_name = inner.split('|').next().unwrap_or(inner).trim();
        if !result.contains(&var_name.to_string()) {
            result.push(var_name.to_string());
        }
    }

    result
}

#[cfg(test)]
mod tests {
    use super::*;
    use tempfile::TempDir;
    use xlex_core::{CellRef, CellValue, Workbook};

    fn default_global() -> GlobalOptions {
        GlobalOptions {
            verbose: false,
            quiet: true,
            format: OutputFormat::Text,
            dry_run: false,
            no_color: true,
            color: false,
            json_errors: false,
            output: None,
        }
    }

    fn create_template_workbook(
        dir: &TempDir,
        name: &str,
        cells: &[(&str, &str)],
    ) -> std::path::PathBuf {
        let path = dir.path().join(name);
        let mut wb = Workbook::new();
        let sheet = wb.get_sheet_mut("Sheet1").unwrap();
        for (cell_ref, value) in cells {
            sheet.set_cell(
                CellRef::parse(cell_ref).unwrap(),
                CellValue::String(value.to_string()),
            );
        }
        wb.save_as(&path).unwrap();
        path
    }

    // ==========================================================================
    // TemplateVars tests
    // ==========================================================================

    #[test]
    fn test_template_vars_new() {
        let vars = TemplateVars::new();
        assert!(vars.get("nonexistent").is_none());
    }

    #[test]
    fn test_template_vars_set_get() {
        let mut vars = TemplateVars::new();
        vars.set("name", "Alice");
        vars.set("age", "30");

        assert_eq!(vars.get("name"), Some("Alice".to_string()));
        assert_eq!(vars.get("age"), Some("30".to_string()));
        assert!(vars.get("unknown").is_none());
    }

    #[test]
    fn test_template_vars_from_json() {
        let json = serde_json::json!({
            "name": "Bob",
            "nested": {
                "value": "inner"
            }
        });
        let vars = TemplateVars::from_json(json);

        assert_eq!(vars.get("name"), Some("Bob".to_string()));
        assert_eq!(vars.get("nested.value"), Some("inner".to_string()));
    }

    #[test]
    fn test_template_vars_nested_access() {
        let json = serde_json::json!({
            "user": {
                "profile": {
                    "name": "Charlie"
                }
            }
        });
        let vars = TemplateVars::from_json(json);

        assert_eq!(vars.get("user.profile.name"), Some("Charlie".to_string()));
    }

    #[test]
    fn test_template_vars_array_access() {
        let json = serde_json::json!({
            "items": ["a", "b", "c"],
            "records": [
                {"id": 1},
                {"id": 2}
            ]
        });
        let vars = TemplateVars::from_json(json);

        let items = vars.get_array("items");
        assert!(items.is_some());
        assert_eq!(items.unwrap().len(), 3);

        let records = vars.get_array("records");
        assert!(records.is_some());
        assert_eq!(records.unwrap().len(), 2);
    }

    #[test]
    fn test_template_vars_get_bool() {
        let json = serde_json::json!({
            "flag_true": true,
            "flag_false": false,
            "str_yes": "yes",
            "str_no": "no",
            "str_1": "1",
            "num_1": 1,
            "num_0": 0,
            "empty": "",
            "null_val": null
        });
        let vars = TemplateVars::from_json(json);

        assert_eq!(vars.get_bool("flag_true"), Some(true));
        assert_eq!(vars.get_bool("flag_false"), Some(false));
        assert_eq!(vars.get_bool("str_yes"), Some(true));
        assert_eq!(vars.get_bool("str_no"), Some(false));
        assert_eq!(vars.get_bool("str_1"), Some(true));
        assert_eq!(vars.get_bool("num_1"), Some(true));
        assert_eq!(vars.get_bool("num_0"), Some(false));
        assert_eq!(vars.get_bool("empty"), Some(false));
        assert_eq!(vars.get_bool("null_val"), Some(false));
    }

    #[test]
    fn test_template_vars_merge_object() {
        let mut vars = TemplateVars::new();
        vars.set("existing", "value");

        let new_obj = serde_json::json!({
            "new_key": "new_value",
            "existing": "overwritten"
        });
        vars.merge_object(&new_obj);

        assert_eq!(vars.get("new_key"), Some("new_value".to_string()));
        assert_eq!(vars.get("existing"), Some("overwritten".to_string()));
    }

    // ==========================================================================
    // find_placeholders tests
    // ==========================================================================

    #[test]
    fn test_find_placeholders_simple() {
        let result = find_placeholders("Hello {{name}}!");
        assert_eq!(result, vec!["name"]);
    }

    #[test]
    fn test_find_placeholders_multiple() {
        let result = find_placeholders("{{first}} and {{second}}");
        assert_eq!(result, vec!["first", "second"]);
    }

    #[test]
    fn test_find_placeholders_with_filters() {
        let result = find_placeholders("{{name|upper}} {{amount|currency}}");
        assert_eq!(result, vec!["name", "amount"]);
    }

    #[test]
    fn test_find_placeholders_no_duplicates() {
        let result = find_placeholders("{{name}} is {{name}}");
        assert_eq!(result, vec!["name"]);
    }

    #[test]
    fn test_find_placeholders_ignores_conditionals() {
        let result = find_placeholders("{{#if show}}content{{/if}}");
        assert!(result.is_empty());
    }

    // ==========================================================================
    // apply_filter tests
    // ==========================================================================

    #[test]
    fn test_filter_upper() {
        assert_eq!(apply_filter("hello", "upper"), "HELLO");
        assert_eq!(apply_filter("hello", "uppercase"), "HELLO");
    }

    #[test]
    fn test_filter_lower() {
        assert_eq!(apply_filter("HELLO", "lower"), "hello");
        assert_eq!(apply_filter("HELLO", "lowercase"), "hello");
    }

    #[test]
    fn test_filter_capitalize() {
        assert_eq!(apply_filter("hello world", "capitalize"), "Hello world");
        assert_eq!(apply_filter("hello world", "title"), "Hello world");
    }

    #[test]
    fn test_filter_trim() {
        assert_eq!(apply_filter("  hello  ", "trim"), "hello");
    }

    #[test]
    fn test_filter_default() {
        assert_eq!(apply_filter("", "default:N/A"), "N/A");
        assert_eq!(apply_filter("value", "default:N/A"), "value");
    }

    #[test]
    fn test_filter_truncate() {
        assert_eq!(apply_filter("hello world", "truncate:5"), "hello...");
        assert_eq!(apply_filter("hi", "truncate:5"), "hi");
    }

    #[test]
    fn test_filter_replace() {
        assert_eq!(
            apply_filter("hello world", "replace:world:universe"),
            "hello universe"
        );
    }

    #[test]
    fn test_filter_currency() {
        assert_eq!(apply_filter("100", "currency"), "$100.00");
        assert_eq!(apply_filter("100", "currency:€"), "€100.00");
        assert_eq!(apply_filter("invalid", "currency"), "invalid");
    }

    #[test]
    fn test_filter_number() {
        assert_eq!(apply_filter("3.14159", "number:2"), "3.14");
        assert_eq!(apply_filter("3.14159", "format_number:3"), "3.142");
    }

    #[test]
    fn test_filter_percent() {
        assert_eq!(apply_filter("0.5", "percent"), "50.0%");
        assert_eq!(apply_filter("0.123", "percent"), "12.3%");
    }

    #[test]
    fn test_filter_abs() {
        assert_eq!(apply_filter("-5", "abs"), "5");
        assert_eq!(apply_filter("5", "abs"), "5");
    }

    #[test]
    fn test_filter_round() {
        assert_eq!(apply_filter("3.567", "round"), "4");
        assert_eq!(apply_filter("3.567", "round:2"), "3.57");
    }

    #[test]
    fn test_filter_unknown() {
        // Unknown filter should return value unchanged
        assert_eq!(apply_filter("value", "unknown_filter"), "value");
    }

    // ==========================================================================
    // evaluate_condition tests
    // ==========================================================================

    #[test]
    fn test_evaluate_condition_equality() {
        let mut vars = TemplateVars::new();
        vars.set("status", "active");

        assert!(evaluate_condition("status == 'active'", &vars));
        assert!(!evaluate_condition("status == 'inactive'", &vars));
    }

    #[test]
    fn test_evaluate_condition_inequality() {
        let mut vars = TemplateVars::new();
        vars.set("status", "active");

        assert!(evaluate_condition("status != 'inactive'", &vars));
        assert!(!evaluate_condition("status != 'active'", &vars));
    }

    #[test]
    fn test_evaluate_condition_numeric_comparisons() {
        let mut vars = TemplateVars::new();
        vars.set("value", "10");

        assert!(evaluate_condition("value > 5", &vars));
        assert!(!evaluate_condition("value > 15", &vars));
        assert!(evaluate_condition("value < 15", &vars));
        assert!(!evaluate_condition("value < 5", &vars));
        assert!(evaluate_condition("value >= 10", &vars));
        assert!(evaluate_condition("value <= 10", &vars));
    }

    #[test]
    fn test_evaluate_condition_boolean() {
        let json = serde_json::json!({
            "show": true,
            "hide": false
        });
        let vars = TemplateVars::from_json(json);

        assert!(evaluate_condition("show", &vars));
        assert!(!evaluate_condition("hide", &vars));
    }

    // ==========================================================================
    // process_template_string tests
    // ==========================================================================

    #[test]
    fn test_process_template_simple() {
        let mut vars = TemplateVars::new();
        vars.set("name", "Alice");

        let result = process_template_string("Hello {{name}}!", &vars);
        assert_eq!(result, "Hello Alice!");
    }

    #[test]
    fn test_process_template_with_filter() {
        let mut vars = TemplateVars::new();
        vars.set("name", "alice");

        let result = process_template_string("Hello {{name|upper}}!", &vars);
        assert_eq!(result, "Hello ALICE!");
    }

    #[test]
    fn test_process_template_conditional() {
        let json = serde_json::json!({
            "show": true,
            "name": "Bob"
        });
        let vars = TemplateVars::from_json(json);

        let result = process_template_string("{{#if show}}Hello {{name}}{{/if}}", &vars);
        assert_eq!(result, "Hello Bob");
    }

    #[test]
    fn test_process_template_conditional_false() {
        let json = serde_json::json!({
            "show": false,
            "name": "Bob"
        });
        let vars = TemplateVars::from_json(json);

        let result = process_template_string("{{#if show}}Hello {{name}}{{/if}}", &vars);
        assert_eq!(result, "");
    }

    #[test]
    fn test_process_template_conditional_else() {
        let json = serde_json::json!({
            "premium": false
        });
        let vars = TemplateVars::from_json(json);

        let result =
            process_template_string("{{#if premium}}Premium User{{else}}Free User{{/if}}", &vars);
        assert_eq!(result, "Free User");
    }

    #[test]
    fn test_process_template_unless() {
        let json = serde_json::json!({
            "hide": false
        });
        let vars = TemplateVars::from_json(json);

        let result = process_template_string("{{#unless hide}}Visible{{/unless}}", &vars);
        assert_eq!(result, "Visible");
    }

    // ==========================================================================
    // value_to_string tests
    // ==========================================================================

    #[test]
    fn test_value_to_string() {
        assert_eq!(value_to_string(&serde_json::json!("hello")), "hello");
        assert_eq!(value_to_string(&serde_json::json!(42)), "42");
        assert_eq!(value_to_string(&serde_json::json!(3.14)), "3.14");
        assert_eq!(value_to_string(&serde_json::json!(true)), "true");
        assert_eq!(value_to_string(&serde_json::json!(null)), "");
    }

    // ==========================================================================
    // yaml_to_json tests
    // ==========================================================================

    #[test]
    fn test_yaml_to_json_string() {
        let yaml: serde_yaml::Value = serde_yaml::from_str("hello").unwrap();
        let json = yaml_to_json(&yaml);
        assert_eq!(json, serde_json::json!("hello"));
    }

    #[test]
    fn test_yaml_to_json_number() {
        let yaml: serde_yaml::Value = serde_yaml::from_str("42").unwrap();
        let json = yaml_to_json(&yaml);
        assert_eq!(json, serde_json::json!(42));
    }

    #[test]
    fn test_yaml_to_json_object() {
        let yaml: serde_yaml::Value = serde_yaml::from_str("name: Alice\nage: 30").unwrap();
        let json = yaml_to_json(&yaml);
        assert_eq!(json, serde_json::json!({"name": "Alice", "age": 30}));
    }

    #[test]
    fn test_yaml_to_json_array() {
        let yaml: serde_yaml::Value = serde_yaml::from_str("- a\n- b\n- c").unwrap();
        let json = yaml_to_json(&yaml);
        assert_eq!(json, serde_json::json!(["a", "b", "c"]));
    }

    // ==========================================================================
    // init() tests
    // ==========================================================================

    #[test]
    fn test_init_report_template() {
        let dir = TempDir::new().unwrap();
        let output = dir.path().join("report.xlsx");

        let result = init(&output, "report", &default_global());
        assert!(result.is_ok());
        assert!(output.exists());

        let wb = Workbook::open(&output).unwrap();
        let sheet = wb.get_sheet("Sheet1").unwrap();
        let cell_a1 = sheet.get_cell(&CellRef::parse("A1").unwrap()).unwrap();
        if let CellValue::String(s) = &cell_a1.value {
            assert!(s.contains("{{title}}"));
        }
    }

    #[test]
    fn test_init_invoice_template() {
        let dir = TempDir::new().unwrap();
        let output = dir.path().join("invoice.xlsx");

        let result = init(&output, "invoice", &default_global());
        assert!(result.is_ok());
        assert!(output.exists());

        let wb = Workbook::open(&output).unwrap();
        let sheet = wb.get_sheet("Sheet1").unwrap();
        let cell_a1 = sheet.get_cell(&CellRef::parse("A1").unwrap()).unwrap();
        if let CellValue::String(s) = &cell_a1.value {
            assert!(s.contains("INVOICE"));
        }
    }

    #[test]
    fn test_init_data_template() {
        let dir = TempDir::new().unwrap();
        let output = dir.path().join("data.xlsx");

        let result = init(&output, "data", &default_global());
        assert!(result.is_ok());
        assert!(output.exists());

        let wb = Workbook::open(&output).unwrap();
        let sheet = wb.get_sheet("Sheet1").unwrap();
        let cell_a1 = sheet.get_cell(&CellRef::parse("A1").unwrap()).unwrap();
        if let CellValue::String(s) = &cell_a1.value {
            assert!(s.contains("{{header1}}"));
        }
    }

    #[test]
    fn test_init_unknown_type() {
        let dir = TempDir::new().unwrap();
        let output = dir.path().join("unknown.xlsx");

        let result = init(&output, "unknown_type", &default_global());
        assert!(result.is_err());
    }

    #[test]
    fn test_init_dry_run() {
        let dir = TempDir::new().unwrap();
        let output = dir.path().join("dryrun.xlsx");

        let global = GlobalOptions {
            dry_run: true,
            ..default_global()
        };

        let result = init(&output, "report", &global);
        assert!(result.is_ok());
        assert!(!output.exists()); // File should not be created in dry-run
    }

    #[test]
    fn test_init_file_exists() {
        let dir = TempDir::new().unwrap();
        let output = dir.path().join("existing.xlsx");

        // Create the file first
        Workbook::new().save_as(&output).unwrap();

        let result = init(&output, "report", &default_global());
        assert!(result.is_err()); // Should fail because file exists
    }

    // ==========================================================================
    // list() tests
    // ==========================================================================

    #[test]
    fn test_list_placeholders() {
        let dir = TempDir::new().unwrap();
        let template = create_template_workbook(
            &dir,
            "template.xlsx",
            &[
                ("A1", "{{name}}"),
                ("A2", "{{email}}"),
                ("B1", "{{name}}"), // duplicate
            ],
        );

        let result = list(&template, &default_global());
        assert!(result.is_ok());
    }

    #[test]
    fn test_list_placeholders_json_format() {
        let dir = TempDir::new().unwrap();
        let template = create_template_workbook(&dir, "template.xlsx", &[("A1", "{{title}}")]);

        let global = GlobalOptions {
            format: OutputFormat::Json,
            ..default_global()
        };

        let result = list(&template, &global);
        assert!(result.is_ok());
    }

    // ==========================================================================
    // validate() tests
    // ==========================================================================

    #[test]
    fn test_validate_generate_schema() {
        let dir = TempDir::new().unwrap();
        let template = create_template_workbook(
            &dir,
            "template.xlsx",
            &[("A1", "{{name}}"), ("A2", "{{email}}")],
        );

        let result = validate(&template, None, true, &default_global());
        assert!(result.is_ok());
    }

    #[test]
    fn test_validate_with_json_vars() {
        let dir = TempDir::new().unwrap();
        let template = create_template_workbook(
            &dir,
            "template.xlsx",
            &[("A1", "{{name}}"), ("A2", "{{email}}")],
        );

        // Create vars file
        let vars_path = dir.path().join("vars.json");
        std::fs::write(
            &vars_path,
            r#"{"name": "Alice", "email": "alice@example.com"}"#,
        )
        .unwrap();

        let result = validate(&template, Some(&vars_path), false, &default_global());
        assert!(result.is_ok());
    }

    #[test]
    fn test_validate_with_yaml_vars() {
        let dir = TempDir::new().unwrap();
        let template = create_template_workbook(&dir, "template.xlsx", &[("A1", "{{name}}")]);

        // Create YAML vars file
        let vars_path = dir.path().join("vars.yaml");
        std::fs::write(&vars_path, "name: Alice\n").unwrap();

        let result = validate(&template, Some(&vars_path), false, &default_global());
        assert!(result.is_ok());
    }

    // ==========================================================================
    // preview() tests
    // ==========================================================================

    #[test]
    fn test_preview_with_defines() {
        let dir = TempDir::new().unwrap();
        let template =
            create_template_workbook(&dir, "template.xlsx", &[("A1", "Hello {{name}}!")]);

        let defines = vec!["name=World".to_string()];
        let result = preview(&template, None, &defines, &default_global());
        assert!(result.is_ok());
    }

    #[test]
    fn test_preview_json_format() {
        let dir = TempDir::new().unwrap();
        let template = create_template_workbook(&dir, "template.xlsx", &[("A1", "{{greeting}}")]);

        let defines = vec!["greeting=Hi".to_string()];
        let global = GlobalOptions {
            format: OutputFormat::Json,
            ..default_global()
        };

        let result = preview(&template, None, &defines, &global);
        assert!(result.is_ok());
    }

    // ==========================================================================
    // create() tests
    // ==========================================================================

    #[test]
    fn test_create_template_from_source() {
        let dir = TempDir::new().unwrap();

        // Create source workbook
        let source = dir.path().join("source.xlsx");
        let mut wb = Workbook::new();
        wb.get_sheet_mut("Sheet1").unwrap().set_cell(
            CellRef::parse("A1").unwrap(),
            CellValue::String("Title".to_string()),
        );
        wb.save_as(&source).unwrap();

        let output = dir.path().join("template.xlsx");
        let placeholders = vec!["A1=title".to_string()];

        let result = create(&source, &output, &placeholders, &default_global());
        assert!(result.is_ok());
        assert!(output.exists());

        // Verify placeholder was added
        let wb = Workbook::open(&output).unwrap();
        let cell = wb
            .get_sheet("Sheet1")
            .unwrap()
            .get_cell(&CellRef::parse("A1").unwrap())
            .unwrap();
        if let CellValue::String(s) = &cell.value {
            assert_eq!(s, "{{title}}");
        }
    }

    #[test]
    fn test_create_with_sheet_reference() {
        let dir = TempDir::new().unwrap();

        let source = dir.path().join("source.xlsx");
        let mut wb = Workbook::new();
        wb.get_sheet_mut("Sheet1").unwrap().set_cell(
            CellRef::parse("B2").unwrap(),
            CellValue::String("Value".to_string()),
        );
        wb.save_as(&source).unwrap();

        let output = dir.path().join("template.xlsx");
        let placeholders = vec!["Sheet1!B2=myvalue".to_string()];

        let result = create(&source, &output, &placeholders, &default_global());
        assert!(result.is_ok());
    }

    #[test]
    fn test_create_dry_run() {
        let dir = TempDir::new().unwrap();

        let source = dir.path().join("source.xlsx");
        Workbook::new().save_as(&source).unwrap();

        let output = dir.path().join("template.xlsx");
        let placeholders = vec!["A1=test".to_string()];

        let global = GlobalOptions {
            dry_run: true,
            ..default_global()
        };

        let result = create(&source, &output, &placeholders, &global);
        assert!(result.is_ok());
        assert!(!output.exists());
    }

    // ==========================================================================
    // apply_single() tests
    // ==========================================================================

    #[test]
    fn test_apply_single() {
        let dir = TempDir::new().unwrap();
        let template = create_template_workbook(
            &dir,
            "template.xlsx",
            &[("A1", "Hello {{name}}!"), ("A2", "Email: {{email}}")],
        );

        let output = dir.path().join("output.xlsx");
        let mut vars = TemplateVars::new();
        vars.set("name", "Alice");
        vars.set("email", "alice@example.com");

        let result = apply_single(&template, &output, &vars, &default_global());
        assert!(result.is_ok());
        assert!(output.exists());

        let wb = Workbook::open(&output).unwrap();
        let sheet = wb.get_sheet("Sheet1").unwrap();

        let cell_a1 = sheet.get_cell(&CellRef::parse("A1").unwrap()).unwrap();
        if let CellValue::String(s) = &cell_a1.value {
            assert_eq!(s, "Hello Alice!");
        }
    }

    #[test]
    fn test_apply_single_with_filters() {
        let dir = TempDir::new().unwrap();
        let template = create_template_workbook(&dir, "template.xlsx", &[("A1", "{{name|upper}}")]);

        let output = dir.path().join("output.xlsx");
        let mut vars = TemplateVars::new();
        vars.set("name", "alice");

        let result = apply_single(&template, &output, &vars, &default_global());
        assert!(result.is_ok());

        let wb = Workbook::open(&output).unwrap();
        let cell = wb
            .get_sheet("Sheet1")
            .unwrap()
            .get_cell(&CellRef::parse("A1").unwrap())
            .unwrap();
        if let CellValue::String(s) = &cell.value {
            assert_eq!(s, "ALICE");
        }
    }

    #[test]
    fn test_apply_single_dry_run() {
        // Note: apply_single does not check dry_run flag
        // The dry_run check is in the apply() dispatcher (before calling apply_single)
        // So this test verifies that apply_single still creates the file
        let dir = TempDir::new().unwrap();
        let template = create_template_workbook(&dir, "template.xlsx", &[("A1", "{{test}}")]);

        let output = dir.path().join("output.xlsx");
        let vars = TemplateVars::new();

        let global = GlobalOptions {
            dry_run: true,
            ..default_global()
        };

        let result = apply_single(&template, &output, &vars, &global);
        assert!(result.is_ok());
        // apply_single ignores dry_run, file is created
        assert!(output.exists());
    }

    // ==========================================================================
    // load_template_vars tests
    // ==========================================================================

    #[test]
    fn test_load_template_vars_json() {
        let dir = TempDir::new().unwrap();
        let vars_path = dir.path().join("vars.json");
        std::fs::write(&vars_path, r#"{"name": "Test", "value": 42}"#).unwrap();

        let vars = load_template_vars(Some(&vars_path), &[]).unwrap();
        assert_eq!(vars.get("name"), Some("Test".to_string()));
        assert_eq!(vars.get("value"), Some("42".to_string()));
    }

    #[test]
    fn test_load_template_vars_yaml() {
        let dir = TempDir::new().unwrap();
        let vars_path = dir.path().join("vars.yaml");
        std::fs::write(&vars_path, "name: YamlTest\ncount: 10").unwrap();

        let vars = load_template_vars(Some(&vars_path), &[]).unwrap();
        assert_eq!(vars.get("name"), Some("YamlTest".to_string()));
        assert_eq!(vars.get("count"), Some("10".to_string()));
    }

    #[test]
    fn test_load_template_vars_with_defines() {
        let vars = load_template_vars(
            None,
            &["key1=value1".to_string(), "key2=value2".to_string()],
        )
        .unwrap();
        assert_eq!(vars.get("key1"), Some("value1".to_string()));
        assert_eq!(vars.get("key2"), Some("value2".to_string()));
    }

    #[test]
    fn test_load_template_vars_defines_override() {
        let dir = TempDir::new().unwrap();
        let vars_path = dir.path().join("vars.json");
        std::fs::write(&vars_path, r#"{"name": "FromFile"}"#).unwrap();

        let vars = load_template_vars(Some(&vars_path), &["name=FromDefine".to_string()]).unwrap();
        assert_eq!(vars.get("name"), Some("FromDefine".to_string()));
    }

    // ==========================================================================
    // process_conditionals tests
    // ==========================================================================

    #[test]
    fn test_process_conditionals_if_true() {
        let json = serde_json::json!({"show": true});
        let vars = TemplateVars::from_json(json);

        let result = process_conditionals("Before {{#if show}}SHOWN{{/if}} After", &vars);
        assert_eq!(result, "Before SHOWN After");
    }

    #[test]
    fn test_process_conditionals_if_false() {
        let json = serde_json::json!({"show": false});
        let vars = TemplateVars::from_json(json);

        let result = process_conditionals("Before {{#if show}}SHOWN{{/if}} After", &vars);
        assert_eq!(result, "Before  After");
    }

    #[test]
    fn test_process_conditionals_if_else() {
        let json = serde_json::json!({"premium": false});
        let vars = TemplateVars::from_json(json);

        let result = process_conditionals("{{#if premium}}Premium{{else}}Basic{{/if}}", &vars);
        assert_eq!(result, "Basic");
    }

    #[test]
    fn test_process_conditionals_unless() {
        let json = serde_json::json!({"hide": false});
        let vars = TemplateVars::from_json(json);

        let result = process_conditionals("{{#unless hide}}Visible{{/unless}}", &vars);
        assert_eq!(result, "Visible");
    }

    #[test]
    fn test_process_conditionals_unless_true() {
        let json = serde_json::json!({"hide": true});
        let vars = TemplateVars::from_json(json);

        let result = process_conditionals("{{#unless hide}}Visible{{/unless}}", &vars);
        assert_eq!(result, "");
    }

    // ==========================================================================
    // resolve_value / resolve_number tests
    // ==========================================================================

    #[test]
    fn test_resolve_value_quoted() {
        let vars = TemplateVars::new();
        assert_eq!(resolve_value("'hello'", &vars), "hello");
        assert_eq!(resolve_value("\"world\"", &vars), "world");
    }

    #[test]
    fn test_resolve_value_variable() {
        let mut vars = TemplateVars::new();
        vars.set("myvar", "myvalue");
        assert_eq!(resolve_value("myvar", &vars), "myvalue");
    }

    #[test]
    fn test_resolve_number() {
        let mut vars = TemplateVars::new();
        vars.set("num", "42");

        assert_eq!(resolve_number("num", &vars), Some(42.0));
        assert_eq!(resolve_number("100", &vars), Some(100.0));
        assert_eq!(resolve_number("3.14", &vars), Some(3.14));
    }

    // ==========================================================================
    // run() integration tests
    // ==========================================================================

    #[test]
    fn test_run_init_command() {
        let dir = TempDir::new().unwrap();
        let output = dir.path().join("test_init.xlsx");

        let args = TemplateArgs {
            command: TemplateCommand::Init {
                output,
                template_type: "data".to_string(),
            },
        };

        let result = run(&args, &default_global());
        assert!(result.is_ok());
    }

    #[test]
    fn test_run_list_command() {
        let dir = TempDir::new().unwrap();
        let template =
            create_template_workbook(&dir, "template.xlsx", &[("A1", "{{placeholder}}")]);

        let args = TemplateArgs {
            command: TemplateCommand::List { template },
        };

        let result = run(&args, &default_global());
        assert!(result.is_ok());
    }

    #[test]
    fn test_run_validate_command() {
        let dir = TempDir::new().unwrap();
        let template = create_template_workbook(&dir, "template.xlsx", &[("A1", "{{test}}")]);

        let args = TemplateArgs {
            command: TemplateCommand::Validate {
                template,
                vars: None,
                schema: false,
            },
        };

        let result = run(&args, &default_global());
        assert!(result.is_ok());
    }

    #[test]
    fn test_run_create_command() {
        let dir = TempDir::new().unwrap();

        let source = dir.path().join("source.xlsx");
        Workbook::new().save_as(&source).unwrap();

        let output = dir.path().join("new_template.xlsx");

        let args = TemplateArgs {
            command: TemplateCommand::Create {
                source,
                output,
                placeholder: vec!["A1=myfield".to_string()],
            },
        };

        let result = run(&args, &default_global());
        assert!(result.is_ok());
    }

    #[test]
    fn test_run_preview_command() {
        let dir = TempDir::new().unwrap();
        let template = create_template_workbook(&dir, "template.xlsx", &[("A1", "{{msg}}")]);

        let args = TemplateArgs {
            command: TemplateCommand::Preview {
                template,
                vars: None,
                define: vec!["msg=Hello".to_string()],
            },
        };

        let result = run(&args, &default_global());
        assert!(result.is_ok());
    }

    #[test]
    fn test_run_apply_command() {
        let dir = TempDir::new().unwrap();
        let template = create_template_workbook(&dir, "template.xlsx", &[("A1", "{{value}}")]);
        let output = dir.path().join("applied.xlsx");

        let args = TemplateArgs {
            command: TemplateCommand::Apply {
                template,
                output,
                vars: None,
                define: vec!["value=123".to_string()],
                per_record: false,
                output_pattern: None,
            },
        };

        let result = run(&args, &default_global());
        assert!(result.is_ok());
    }

    #[test]
    fn test_apply_single_json_output() {
        let dir = TempDir::new().unwrap();
        let template = create_template_workbook(&dir, "template.xlsx", &[("A1", "{{value}}")]);
        let output = dir.path().join("applied.xlsx");

        let mut global = default_global();
        global.format = OutputFormat::Json;
        global.quiet = false;

        let vars = TemplateVars::new();
        let result = apply_single(&template, &output, &vars, &global);
        assert!(result.is_ok());
    }

    #[test]
    fn test_list_json_output() {
        let dir = TempDir::new().unwrap();
        let template = create_template_workbook(
            &dir,
            "template.xlsx",
            &[("A1", "{{name}}"), ("B1", "{{age}}")],
        );

        let mut global = default_global();
        global.format = OutputFormat::Json;

        let result = list(&template, &global);
        assert!(result.is_ok());
    }

    #[test]
    fn test_validate_with_missing_vars() {
        let dir = TempDir::new().unwrap();
        let template = create_template_workbook(&dir, "template.xlsx", &[("A1", "{{name}}")]);

        let result = validate(&template, None, false, &default_global());
        // Should show warnings about missing variables but not error
        assert!(result.is_ok());
    }

    #[test]
    fn test_create_with_multiple_placeholders() {
        let dir = TempDir::new().unwrap();

        let source = dir.path().join("source.xlsx");
        {
            let mut wb = Workbook::new();
            wb.set_cell(
                "Sheet1",
                CellRef::new(1, 1),
                CellValue::String("Value1".to_string()),
            )
            .unwrap();
            wb.set_cell(
                "Sheet1",
                CellRef::new(2, 1),
                CellValue::String("Value2".to_string()),
            )
            .unwrap();
            wb.save_as(&source).unwrap();
        }

        let output = dir.path().join("new_template.xlsx");

        let result = create(
            &source,
            &output,
            &["A1=field1".to_string(), "B1=field2".to_string()],
            &default_global(),
        );
        assert!(result.is_ok());
    }

    #[test]
    fn test_preview_with_filters() {
        let dir = TempDir::new().unwrap();
        let template = create_template_workbook(&dir, "template.xlsx", &[("A1", "{{name|upper}}")]);

        let result = preview(
            &template,
            None,
            &["name=test".to_string()],
            &default_global(),
        );
        assert!(result.is_ok());
    }

    #[test]
    fn test_filter_date() {
        // Date filter formats dates
        let result = apply_filter("2024-01-15", "date");
        assert_eq!(result, "2024-01-15");
    }

    #[test]
    fn test_filter_now() {
        // Now filter returns current date
        let result = apply_filter("", "now");
        // Result should be a date string in YYYY-MM-DD format
        assert!(result.len() >= 10);
    }

    #[test]
    fn test_template_vars_get_method() {
        let mut vars = TemplateVars::new();
        vars.set("key", "value");
        assert!(vars.get("key").is_some());
        assert!(vars.get("nonexistent").is_none());
    }

    #[test]
    fn test_apply_per_record_no_records() {
        let dir = TempDir::new().unwrap();
        let template = create_template_workbook(&dir, "template.xlsx", &[("A1", "{{value}}")]);
        let output = dir.path().join("output.xlsx");

        // vars without records array should fail
        let vars = TemplateVars::new();
        let result = apply_per_record(&template, &output, None, &vars, &default_global());
        assert!(result.is_err());
    }
}
