//! Workbook operations.

use anyhow::Result;
use clap::Parser;
use colored::Colorize;

use xlex_core::Workbook;

use super::{GlobalOptions, OutputFormat};

/// Arguments for the info command.
#[derive(Parser)]
pub struct InfoArgs {
    /// Path to the xlsx file
    pub file: std::path::PathBuf,
}

/// Display workbook information.
pub fn info(args: &InfoArgs, global: &GlobalOptions) -> Result<()> {
    let workbook = Workbook::open(&args.file)?;
    let props = workbook.properties();

    if global.format == OutputFormat::Json {
        let json = serde_json::json!({
            "file": args.file.display().to_string(),
            "sheets": workbook.sheet_names(),
            "sheetCount": workbook.sheet_count(),
            "properties": {
                "title": props.title,
                "subject": props.subject,
                "creator": props.creator,
                "keywords": props.keywords,
                "description": props.description,
                "lastModifiedBy": props.last_modified_by,
            }
        });
        println!("{}", serde_json::to_string_pretty(&json)?);
    } else {
        println!("{}: {}", "File".bold(), args.file.display());
        println!("{}: {}", "Sheets".bold(), workbook.sheet_count());

        println!("\n{}:", "Sheet Names".bold());
        for (i, name) in workbook.sheet_names().iter().enumerate() {
            let visibility = workbook.get_sheet_visibility(name).unwrap_or_default();
            let vis_str = if visibility.is_hidden() {
                " (hidden)".dimmed().to_string()
            } else {
                String::new()
            };
            println!("  {}. {}{}", i + 1, name, vis_str);
        }

        if props.title.is_some() || props.creator.is_some() || props.subject.is_some() {
            println!("\n{}:", "Properties".bold());
            if let Some(ref title) = props.title {
                println!("  {}: {}", "Title".cyan(), title);
            }
            if let Some(ref creator) = props.creator {
                println!("  {}: {}", "Creator".cyan(), creator);
            }
            if let Some(ref subject) = props.subject {
                println!("  {}: {}", "Subject".cyan(), subject);
            }
            if let Some(ref description) = props.description {
                println!("  {}: {}", "Description".cyan(), description);
            }
            if let Some(ref keywords) = props.keywords {
                println!("  {}: {}", "Keywords".cyan(), keywords);
            }
        }
    }

    Ok(())
}

/// Arguments for the validate command.
#[derive(Parser)]
pub struct ValidateArgs {
    /// Path to the xlsx file
    pub file: std::path::PathBuf,
}

/// Validate workbook structure.
pub fn validate(args: &ValidateArgs, global: &GlobalOptions) -> Result<()> {
    // Try to open and parse the workbook
    match Workbook::open(&args.file) {
        Ok(_workbook) => {
            if global.format == OutputFormat::Json {
                let json = serde_json::json!({
                    "valid": true,
                    "file": args.file.display().to_string(),
                });
                println!("{}", serde_json::to_string_pretty(&json)?);
            } else if !global.quiet {
                println!("{}: {}", "✓".green(), args.file.display());
                println!("  Workbook is valid");
            }
            Ok(())
        }
        Err(e) => {
            if global.format == OutputFormat::Json {
                let json = serde_json::json!({
                    "valid": false,
                    "file": args.file.display().to_string(),
                    "error": e.to_string(),
                });
                println!("{}", serde_json::to_string_pretty(&json)?);
            }
            Err(e.into())
        }
    }
}

/// Arguments for the clone command.
#[derive(Parser)]
pub struct CloneArgs {
    /// Source xlsx file
    pub source: std::path::PathBuf,
    /// Destination xlsx file
    pub dest: std::path::PathBuf,
    /// Overwrite existing file
    #[arg(long, short = 'F')]
    pub force: bool,
}

/// Clone a workbook.
pub fn clone(args: &CloneArgs, global: &GlobalOptions) -> Result<()> {
    // Check if dest exists
    if args.dest.exists() && !args.force {
        return Err(xlex_core::XlexError::FileExists {
            path: args.dest.clone(),
        }
        .into());
    }

    if global.dry_run {
        println!(
            "Would copy {} to {}",
            args.source.display(),
            args.dest.display()
        );
        return Ok(());
    }

    std::fs::copy(&args.source, &args.dest)?;

    if !global.quiet {
        if global.format == OutputFormat::Json {
            let json = serde_json::json!({
                "source": args.source.display().to_string(),
                "destination": args.dest.display().to_string(),
            });
            println!("{}", serde_json::to_string_pretty(&json)?);
        } else {
            println!(
                "Cloned {} to {}",
                args.source.display().to_string().cyan(),
                args.dest.display().to_string().green()
            );
        }
    }

    Ok(())
}

/// Arguments for the create command.
#[derive(Parser)]
pub struct CreateArgs {
    /// Path to create the xlsx file
    pub file: std::path::PathBuf,
    /// Name of the initial sheet
    #[arg(long, short = 's', default_value = "Sheet1")]
    pub sheet: String,
    /// Create multiple sheets (comma-separated)
    #[arg(long)]
    pub sheets: Option<String>,
    /// Overwrite existing file
    #[arg(long, short = 'F')]
    pub force: bool,
}

/// Create a new workbook.
pub fn create(args: &CreateArgs, global: &GlobalOptions) -> Result<()> {
    // Check if file exists
    if args.file.exists() && !args.force {
        return Err(xlex_core::XlexError::FileExists {
            path: args.file.clone(),
        }
        .into());
    }

    // Determine sheets to create
    let sheet_names: Vec<&str> = if let Some(ref sheets) = args.sheets {
        sheets.split(',').map(|s| s.trim()).collect()
    } else {
        vec![args.sheet.as_str()]
    };

    if global.dry_run {
        println!(
            "Would create {} with sheets: {:?}",
            args.file.display(),
            sheet_names
        );
        return Ok(());
    }

    // Create workbook
    let workbook = Workbook::with_sheets(&sheet_names);
    workbook.save_as(&args.file)?;

    if !global.quiet {
        if global.format == OutputFormat::Json {
            let json = serde_json::json!({
                "file": args.file.display().to_string(),
                "sheets": sheet_names,
            });
            println!("{}", serde_json::to_string_pretty(&json)?);
        } else {
            println!(
                "Created {} with {} sheet(s)",
                args.file.display().to_string().green(),
                sheet_names.len()
            );
        }
    }

    Ok(())
}

/// Arguments for the props command.
#[derive(Parser)]
pub struct PropsArgs {
    #[command(subcommand)]
    pub command: PropsCommand,
}

#[derive(clap::Subcommand)]
pub enum PropsCommand {
    /// Get workbook properties
    Get {
        /// Path to the xlsx file
        file: std::path::PathBuf,
        /// Specific property to get
        property: Option<String>,
    },
    /// Set workbook properties
    Set {
        /// Path to the xlsx file
        file: std::path::PathBuf,
        /// Property name
        property: String,
        /// Property value
        value: String,
    },
}

/// Get or set workbook properties.
pub fn props(args: &PropsArgs, global: &GlobalOptions) -> Result<()> {
    match &args.command {
        PropsCommand::Get { file, property } => {
            let workbook = Workbook::open(file)?;
            let props = workbook.properties();

            if let Some(prop_name) = property {
                let value = match prop_name.as_str() {
                    "title" => props.title.clone(),
                    "subject" => props.subject.clone(),
                    "creator" => props.creator.clone(),
                    "keywords" => props.keywords.clone(),
                    "description" => props.description.clone(),
                    "lastModifiedBy" => props.last_modified_by.clone(),
                    "category" => props.category.clone(),
                    _ => None,
                };

                if global.format == OutputFormat::Json {
                    let json = serde_json::json!({
                        "property": prop_name,
                        "value": value,
                    });
                    println!("{}", serde_json::to_string_pretty(&json)?);
                } else {
                    println!("{}", value.unwrap_or_default());
                }
            } else {
                // Show all properties
                if global.format == OutputFormat::Json {
                    let json = serde_json::json!({
                        "title": props.title,
                        "subject": props.subject,
                        "creator": props.creator,
                        "keywords": props.keywords,
                        "description": props.description,
                        "lastModifiedBy": props.last_modified_by,
                        "category": props.category,
                    });
                    println!("{}", serde_json::to_string_pretty(&json)?);
                } else {
                    if let Some(ref v) = props.title {
                        println!("{}: {}", "title".cyan(), v);
                    }
                    if let Some(ref v) = props.subject {
                        println!("{}: {}", "subject".cyan(), v);
                    }
                    if let Some(ref v) = props.creator {
                        println!("{}: {}", "creator".cyan(), v);
                    }
                    if let Some(ref v) = props.keywords {
                        println!("{}: {}", "keywords".cyan(), v);
                    }
                    if let Some(ref v) = props.description {
                        println!("{}: {}", "description".cyan(), v);
                    }
                    if let Some(ref v) = props.last_modified_by {
                        println!("{}: {}", "lastModifiedBy".cyan(), v);
                    }
                    if let Some(ref v) = props.category {
                        println!("{}: {}", "category".cyan(), v);
                    }
                }
            }
        }
        PropsCommand::Set {
            file,
            property,
            value,
        } => {
            if global.dry_run {
                println!(
                    "Would set {} to '{}' in {}",
                    property,
                    value,
                    file.display()
                );
                return Ok(());
            }

            let mut workbook = Workbook::open(file)?;
            let props = workbook.properties_mut();

            match property.as_str() {
                "title" => props.title = Some(value.clone()),
                "subject" => props.subject = Some(value.clone()),
                "creator" => props.creator = Some(value.clone()),
                "keywords" => props.keywords = Some(value.clone()),
                "description" => props.description = Some(value.clone()),
                "lastModifiedBy" => props.last_modified_by = Some(value.clone()),
                "category" => props.category = Some(value.clone()),
                _ => anyhow::bail!("Unknown property: {}", property),
            }

            workbook.save()?;

            if !global.quiet {
                println!("Set {} to '{}'", property.cyan(), value.green());
            }
        }
    }

    Ok(())
}

/// Arguments for the stats command.
#[derive(Parser)]
pub struct StatsArgs {
    /// Path to the xlsx file
    pub file: std::path::PathBuf,
}

/// Display workbook statistics.
pub fn stats(args: &StatsArgs, global: &GlobalOptions) -> Result<()> {
    let workbook = Workbook::open(&args.file)?;
    let stats = workbook.stats();

    if global.format == OutputFormat::Json {
        println!("{}", serde_json::to_string_pretty(&stats)?);
    } else {
        println!("{}:", "Workbook Statistics".bold());
        println!("  {}: {}", "Sheets".cyan(), stats.sheet_count);
        println!("  {}: {}", "Total Cells".cyan(), stats.total_cells);
        println!("  {}: {}", "Formulas".cyan(), stats.formula_count);
        println!("  {}: {}", "Styles".cyan(), stats.style_count);
        println!("  {}: {}", "Shared Strings".cyan(), stats.string_count);
        if stats.file_size > 0 {
            println!("  {}: {} bytes", "File Size".cyan(), stats.file_size);
        }
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
            format: OutputFormat::Text,
            no_color: true,
            color: false,
            json_errors: false,
            dry_run: false,
            output: None,
        }
    }

    #[test]
    fn test_create_workbook() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = temp_dir.path().join("test.xlsx");

        let args = CreateArgs {
            file: file_path.clone(),
            sheet: "Sheet1".to_string(),
            sheets: None,
            force: false,
        };

        let result = create(&args, &default_global());
        assert!(result.is_ok());
        assert!(file_path.exists());
    }

    #[test]
    fn test_create_workbook_with_multiple_sheets() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = temp_dir.path().join("test_multi.xlsx");

        let args = CreateArgs {
            file: file_path.clone(),
            sheet: "Sheet1".to_string(),
            sheets: Some("Data,Summary,Config".to_string()),
            force: false,
        };

        let result = create(&args, &default_global());
        assert!(result.is_ok());

        // Verify sheets were created
        let wb = Workbook::open(&file_path).unwrap();
        assert_eq!(wb.sheet_count(), 3);
    }

    #[test]
    fn test_create_workbook_exists_no_force() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = temp_dir.path().join("existing.xlsx");

        // Create first
        let args = CreateArgs {
            file: file_path.clone(),
            sheet: "Sheet1".to_string(),
            sheets: None,
            force: false,
        };
        create(&args, &default_global()).unwrap();

        // Try to create again without force
        let result = create(&args, &default_global());
        assert!(result.is_err());
    }

    #[test]
    fn test_create_workbook_exists_with_force() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = temp_dir.path().join("overwrite.xlsx");

        // Create first
        let args = CreateArgs {
            file: file_path.clone(),
            sheet: "Sheet1".to_string(),
            sheets: None,
            force: false,
        };
        create(&args, &default_global()).unwrap();

        // Create again with force
        let args = CreateArgs {
            file: file_path.clone(),
            sheet: "NewSheet".to_string(),
            sheets: None,
            force: true,
        };
        let result = create(&args, &default_global());
        assert!(result.is_ok());
    }

    #[test]
    fn test_create_dry_run() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = temp_dir.path().join("dry_run.xlsx");

        let mut global = default_global();
        global.dry_run = true;

        let args = CreateArgs {
            file: file_path.clone(),
            sheet: "Sheet1".to_string(),
            sheets: None,
            force: false,
        };

        let result = create(&args, &global);
        assert!(result.is_ok());
        // File should not be created
        assert!(!file_path.exists());
    }

    #[test]
    fn test_validate_valid_workbook() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = temp_dir.path().join("valid.xlsx");

        // Create a valid workbook
        let wb = Workbook::new();
        wb.save_as(&file_path).unwrap();

        let args = ValidateArgs { file: file_path };

        let result = validate(&args, &default_global());
        assert!(result.is_ok());
    }

    #[test]
    fn test_validate_invalid_file() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = temp_dir.path().join("invalid.xlsx");

        // Create an invalid file
        std::fs::write(&file_path, "not a valid xlsx").unwrap();

        let args = ValidateArgs { file: file_path };

        let result = validate(&args, &default_global());
        assert!(result.is_err());
    }

    #[test]
    fn test_validate_json_output() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = temp_dir.path().join("valid_json.xlsx");

        let wb = Workbook::new();
        wb.save_as(&file_path).unwrap();

        let mut global = default_global();
        global.format = OutputFormat::Json;

        let args = ValidateArgs { file: file_path };

        let result = validate(&args, &global);
        assert!(result.is_ok());
    }

    #[test]
    fn test_clone_workbook() {
        let temp_dir = TempDir::new().unwrap();
        let source_path = temp_dir.path().join("source.xlsx");
        let dest_path = temp_dir.path().join("dest.xlsx");

        // Create source
        let wb = Workbook::new();
        wb.save_as(&source_path).unwrap();

        let args = CloneArgs {
            source: source_path.clone(),
            dest: dest_path.clone(),
            force: false,
        };

        let result = clone(&args, &default_global());
        assert!(result.is_ok());
        assert!(dest_path.exists());
    }

    #[test]
    fn test_clone_dest_exists_no_force() {
        let temp_dir = TempDir::new().unwrap();
        let source_path = temp_dir.path().join("source.xlsx");
        let dest_path = temp_dir.path().join("dest.xlsx");

        // Create both files
        let wb = Workbook::new();
        wb.save_as(&source_path).unwrap();
        wb.save_as(&dest_path).unwrap();

        let args = CloneArgs {
            source: source_path,
            dest: dest_path,
            force: false,
        };

        let result = clone(&args, &default_global());
        assert!(result.is_err());
    }

    #[test]
    fn test_clone_dry_run() {
        let temp_dir = TempDir::new().unwrap();
        let source_path = temp_dir.path().join("source.xlsx");
        let dest_path = temp_dir.path().join("dest_dry.xlsx");

        let wb = Workbook::new();
        wb.save_as(&source_path).unwrap();

        let mut global = default_global();
        global.dry_run = true;

        let args = CloneArgs {
            source: source_path,
            dest: dest_path.clone(),
            force: false,
        };

        let result = clone(&args, &global);
        assert!(result.is_ok());
        assert!(!dest_path.exists());
    }

    #[test]
    fn test_info() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = temp_dir.path().join("info.xlsx");

        let wb = Workbook::new();
        wb.save_as(&file_path).unwrap();

        let args = InfoArgs { file: file_path };

        let result = info(&args, &default_global());
        assert!(result.is_ok());
    }

    #[test]
    fn test_info_json_output() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = temp_dir.path().join("info_json.xlsx");

        let wb = Workbook::new();
        wb.save_as(&file_path).unwrap();

        let mut global = default_global();
        global.format = OutputFormat::Json;

        let args = InfoArgs { file: file_path };

        let result = info(&args, &global);
        assert!(result.is_ok());
    }

    #[test]
    fn test_stats() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = temp_dir.path().join("stats.xlsx");

        let wb = Workbook::new();
        wb.save_as(&file_path).unwrap();

        let args = StatsArgs { file: file_path };

        let result = stats(&args, &default_global());
        assert!(result.is_ok());
    }

    #[test]
    fn test_stats_json_output() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = temp_dir.path().join("stats_json.xlsx");

        let wb = Workbook::new();
        wb.save_as(&file_path).unwrap();

        let mut global = default_global();
        global.format = OutputFormat::Json;

        let args = StatsArgs { file: file_path };

        let result = stats(&args, &global);
        assert!(result.is_ok());
    }
}
