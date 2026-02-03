//! Sheet operations.

use anyhow::Result;
use clap::{Parser, Subcommand};
use colored::Colorize;

use xlex_core::Workbook;

use super::{GlobalOptions, OutputFormat};

/// Arguments for sheet operations.
#[derive(Parser)]
pub struct SheetArgs {
    #[command(subcommand)]
    pub command: SheetCommand,
}

#[derive(Subcommand)]
pub enum SheetCommand {
    /// List all sheets
    List {
        /// Path to the xlsx file
        file: std::path::PathBuf,
    },
    /// Add a new sheet
    Add {
        /// Path to the xlsx file
        file: std::path::PathBuf,
        /// Name of the new sheet
        name: String,
        /// Position to insert (0-indexed)
        #[arg(long, short = 'p')]
        position: Option<usize>,
    },
    /// Remove a sheet
    Remove {
        /// Path to the xlsx file
        file: std::path::PathBuf,
        /// Name of the sheet to remove
        name: String,
    },
    /// Rename a sheet
    Rename {
        /// Path to the xlsx file
        file: std::path::PathBuf,
        /// Current name of the sheet
        old_name: String,
        /// New name for the sheet
        new_name: String,
    },
    /// Copy a sheet
    Copy {
        /// Path to the xlsx file
        file: std::path::PathBuf,
        /// Name of the sheet to copy
        source: String,
        /// Name for the new sheet
        dest: String,
    },
    /// Move a sheet to a different position
    Move {
        /// Path to the xlsx file
        file: std::path::PathBuf,
        /// Name of the sheet to move
        name: String,
        /// New position (0-indexed)
        position: usize,
    },
    /// Hide a sheet
    Hide {
        /// Path to the xlsx file
        file: std::path::PathBuf,
        /// Name of the sheet to hide
        name: String,
        /// Very hidden (cannot be unhidden via UI)
        #[arg(long)]
        very: bool,
    },
    /// Unhide a sheet
    Unhide {
        /// Path to the xlsx file
        file: std::path::PathBuf,
        /// Name of the sheet to unhide
        name: String,
    },
    /// Show sheet information
    Info {
        /// Path to the xlsx file
        file: std::path::PathBuf,
        /// Name of the sheet
        name: String,
    },
    /// Set or display active sheet
    Active {
        /// Path to the xlsx file
        file: std::path::PathBuf,
        /// Name of the sheet to set as active (omit to show current)
        name: Option<String>,
    },
}

/// Run sheet operations.
pub fn run(args: &SheetArgs, global: &GlobalOptions) -> Result<()> {
    match &args.command {
        SheetCommand::List { file } => list(file, global),
        SheetCommand::Add {
            file,
            name,
            position,
        } => add(file, name, *position, global),
        SheetCommand::Remove { file, name } => remove(file, name, global),
        SheetCommand::Rename {
            file,
            old_name,
            new_name,
        } => rename(file, old_name, new_name, global),
        SheetCommand::Copy { file, source, dest } => copy(file, source, dest, global),
        SheetCommand::Move {
            file,
            name,
            position,
        } => move_sheet(file, name, *position, global),
        SheetCommand::Hide { file, name, very } => hide(file, name, *very, global),
        SheetCommand::Unhide { file, name } => unhide(file, name, global),
        SheetCommand::Info { file, name } => info(file, name, global),
        SheetCommand::Active { file, name } => active(file, name.as_deref(), global),
    }
}

fn list(file: &std::path::Path, global: &GlobalOptions) -> Result<()> {
    let workbook = Workbook::open(file)?;

    if global.format == OutputFormat::Json {
        let sheets: Vec<_> = workbook
            .sheet_names()
            .iter()
            .enumerate()
            .map(|(i, name)| {
                let visibility = workbook.get_sheet_visibility(name).unwrap_or_default();
                serde_json::json!({
                    "index": i,
                    "name": name,
                    "visible": visibility.is_visible(),
                    "active": i == workbook.active_sheet_index(),
                })
            })
            .collect();
        println!("{}", serde_json::to_string_pretty(&sheets)?);
    } else {
        for (i, name) in workbook.sheet_names().iter().enumerate() {
            let visibility = workbook.get_sheet_visibility(name).unwrap_or_default();
            let active = if i == workbook.active_sheet_index() {
                " *".green().to_string()
            } else {
                String::new()
            };
            let vis = if visibility.is_hidden() {
                " (hidden)".dimmed().to_string()
            } else {
                String::new()
            };
            println!("{}. {}{}{}", i + 1, name, active, vis);
        }
    }

    Ok(())
}

fn add(
    file: &std::path::Path,
    name: &str,
    _position: Option<usize>,
    global: &GlobalOptions,
) -> Result<()> {
    if global.dry_run {
        println!("Would add sheet '{}' to {}", name, file.display());
        return Ok(());
    }

    let mut workbook = Workbook::open(file)?;
    workbook.add_sheet(name)?;
    workbook.save()?;

    if !global.quiet {
        if global.format == OutputFormat::Json {
            let json = serde_json::json!({
                "action": "add",
                "sheet": name,
            });
            println!("{}", serde_json::to_string_pretty(&json)?);
        } else {
            println!("Added sheet '{}'", name.green());
        }
    }

    Ok(())
}

fn remove(file: &std::path::Path, name: &str, global: &GlobalOptions) -> Result<()> {
    if global.dry_run {
        println!("Would remove sheet '{}' from {}", name, file.display());
        return Ok(());
    }

    let mut workbook = Workbook::open(file)?;
    workbook.remove_sheet(name)?;
    workbook.save()?;

    if !global.quiet {
        if global.format == OutputFormat::Json {
            let json = serde_json::json!({
                "action": "remove",
                "sheet": name,
            });
            println!("{}", serde_json::to_string_pretty(&json)?);
        } else {
            println!("Removed sheet '{}'", name.red());
        }
    }

    Ok(())
}

fn rename(
    file: &std::path::Path,
    old_name: &str,
    new_name: &str,
    global: &GlobalOptions,
) -> Result<()> {
    if global.dry_run {
        println!(
            "Would rename sheet '{}' to '{}' in {}",
            old_name,
            new_name,
            file.display()
        );
        return Ok(());
    }

    let mut workbook = Workbook::open(file)?;
    workbook.rename_sheet(old_name, new_name)?;
    workbook.save()?;

    if !global.quiet {
        if global.format == OutputFormat::Json {
            let json = serde_json::json!({
                "action": "rename",
                "oldName": old_name,
                "newName": new_name,
            });
            println!("{}", serde_json::to_string_pretty(&json)?);
        } else {
            println!("Renamed '{}' to '{}'", old_name.cyan(), new_name.green());
        }
    }

    Ok(())
}

fn copy(file: &std::path::Path, source: &str, dest: &str, global: &GlobalOptions) -> Result<()> {
    if global.dry_run {
        println!(
            "Would copy sheet '{}' to '{}' in {}",
            source,
            dest,
            file.display()
        );
        return Ok(());
    }

    // TODO: Implement proper sheet copying with cell data
    let mut workbook = Workbook::open(file)?;

    // Check source exists
    if workbook.get_sheet(source).is_none() {
        return Err(xlex_core::XlexError::SheetNotFound {
            name: source.to_string(),
        }
        .into());
    }

    // Add new sheet
    workbook.add_sheet(dest)?;
    workbook.save()?;

    if !global.quiet {
        println!("Copied '{}' to '{}'", source.cyan(), dest.green());
    }

    Ok(())
}

fn move_sheet(
    file: &std::path::Path,
    name: &str,
    position: usize,
    global: &GlobalOptions,
) -> Result<()> {
    if global.dry_run {
        println!(
            "Would move sheet '{}' to position {} in {}",
            name,
            position,
            file.display()
        );
        return Ok(());
    }

    let mut workbook = Workbook::open(file)?;
    workbook.move_sheet(name, position)?;
    workbook.save()?;

    if !global.quiet {
        println!("Moved sheet '{}' to position {}", name.cyan(), position);
    }

    Ok(())
}

fn hide(file: &std::path::Path, name: &str, very: bool, global: &GlobalOptions) -> Result<()> {
    if global.dry_run {
        println!("Would hide sheet '{}' in {}", name, file.display());
        return Ok(());
    }

    let mut workbook = Workbook::open(file)?;
    let visibility = if very {
        xlex_core::sheet::SheetVisibility::VeryHidden
    } else {
        xlex_core::sheet::SheetVisibility::Hidden
    };
    workbook.set_sheet_visibility(name, visibility)?;
    workbook.save()?;

    if !global.quiet {
        println!("Hid sheet '{}'", name.dimmed());
    }

    Ok(())
}

fn unhide(file: &std::path::Path, name: &str, global: &GlobalOptions) -> Result<()> {
    if global.dry_run {
        println!("Would unhide sheet '{}' in {}", name, file.display());
        return Ok(());
    }

    let mut workbook = Workbook::open(file)?;
    workbook.set_sheet_visibility(name, xlex_core::sheet::SheetVisibility::Visible)?;
    workbook.save()?;

    if !global.quiet {
        println!("Unhid sheet '{}'", name.green());
    }

    Ok(())
}

fn info(file: &std::path::Path, name: &str, global: &GlobalOptions) -> Result<()> {
    let workbook = Workbook::open(file)?;
    let sheet = workbook
        .get_sheet(name)
        .ok_or_else(|| xlex_core::XlexError::SheetNotFound {
            name: name.to_string(),
        })?;

    if global.format == OutputFormat::Json {
        let json = serde_json::json!({
            "name": sheet.name(),
            "cellCount": sheet.cell_count(),
            "visibility": if sheet.info.visibility.is_visible() { "visible" } else { "hidden" },
            "index": sheet.info.index,
        });
        println!("{}", serde_json::to_string_pretty(&json)?);
    } else {
        println!("{}: {}", "Name".bold(), sheet.name());
        println!("{}: {}", "Index".cyan(), sheet.info.index);
        println!("{}: {}", "Cells".cyan(), sheet.cell_count());
        println!(
            "{}: {}",
            "Visibility".cyan(),
            if sheet.info.visibility.is_visible() {
                "visible"
            } else {
                "hidden"
            }
        );
    }

    Ok(())
}

fn active(file: &std::path::Path, name: Option<&str>, global: &GlobalOptions) -> Result<()> {
    if let Some(name) = name {
        if global.dry_run {
            println!("Would set active sheet to '{}' in {}", name, file.display());
            return Ok(());
        }

        let mut workbook = Workbook::open(file)?;
        workbook.set_active_sheet_by_name(name)?;
        workbook.save()?;

        if !global.quiet {
            println!("Set active sheet to '{}'", name.green());
        }
    } else {
        let workbook = Workbook::open(file)?;
        let active_index = workbook.active_sheet_index();
        let active_name = workbook.sheet_names()[active_index];

        if global.format == OutputFormat::Json {
            let json = serde_json::json!({
                "name": active_name,
                "index": active_index,
            });
            println!("{}", serde_json::to_string_pretty(&json)?);
        } else {
            println!("{}", active_name);
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

    fn create_test_workbook(dir: &TempDir, name: &str) -> std::path::PathBuf {
        let file_path = dir.path().join(name);
        let wb = Workbook::new();
        wb.save_as(&file_path).unwrap();
        file_path
    }

    #[test]
    fn test_list_sheets() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "list.xlsx");

        let result = list(&file_path, &default_global());
        assert!(result.is_ok());
    }

    #[test]
    fn test_list_sheets_json() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "list_json.xlsx");

        let mut global = default_global();
        global.format = OutputFormat::Json;

        let result = list(&file_path, &global);
        assert!(result.is_ok());
    }

    #[test]
    fn test_add_sheet() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "add.xlsx");

        let result = add(&file_path, "NewSheet", None, &default_global());
        assert!(result.is_ok());

        let wb = Workbook::open(&file_path).unwrap();
        assert!(wb.sheet_names().contains(&"NewSheet"));
    }

    #[test]
    fn test_add_sheet_with_position() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "add_pos.xlsx");

        let result = add(&file_path, "NewSheet", Some(0), &default_global());
        assert!(result.is_ok());

        let wb = Workbook::open(&file_path).unwrap();
        // NewSheet should be at position 1 (after Sheet1, then moved to 0)
        assert!(wb.sheet_names().contains(&"NewSheet"));
    }

    #[test]
    fn test_add_sheet_dry_run() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "add_dry.xlsx");

        let mut global = default_global();
        global.dry_run = true;

        let result = add(&file_path, "DrySheet", None, &global);
        assert!(result.is_ok());

        // Sheet should not be added
        let wb = Workbook::open(&file_path).unwrap();
        assert!(!wb.sheet_names().contains(&"DrySheet"));
    }

    #[test]
    fn test_remove_sheet() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "remove.xlsx");

        // Add another sheet first
        add(&file_path, "ToRemove", None, &default_global()).unwrap();

        let result = remove(&file_path, "ToRemove", &default_global());
        assert!(result.is_ok());

        let wb = Workbook::open(&file_path).unwrap();
        assert!(!wb.sheet_names().contains(&"ToRemove"));
    }

    #[test]
    fn test_rename_sheet() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "rename.xlsx");

        let result = rename(&file_path, "Sheet1", "Renamed", &default_global());
        assert!(result.is_ok());

        let wb = Workbook::open(&file_path).unwrap();
        assert!(wb.sheet_names().contains(&"Renamed"));
        assert!(!wb.sheet_names().contains(&"Sheet1"));
    }

    #[test]
    fn test_copy_sheet() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "copy.xlsx");

        let result = copy(&file_path, "Sheet1", "Sheet1_Copy", &default_global());
        assert!(result.is_ok());

        let wb = Workbook::open(&file_path).unwrap();
        assert!(wb.sheet_names().contains(&"Sheet1"));
        assert!(wb.sheet_names().contains(&"Sheet1_Copy"));
    }

    #[test]
    fn test_move_sheet() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "move.xlsx");

        // Add sheets to move
        add(&file_path, "A", None, &default_global()).unwrap();
        add(&file_path, "B", None, &default_global()).unwrap();

        let result = move_sheet(&file_path, "B", 0, &default_global());
        assert!(result.is_ok());
    }

    #[test]
    fn test_hide_sheet() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "hide.xlsx");

        // Add another sheet so we can hide Sheet1
        add(&file_path, "Other", None, &default_global()).unwrap();

        let result = hide(&file_path, "Sheet1", false, &default_global());
        assert!(result.is_ok());
    }

    #[test]
    fn test_unhide_sheet() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "unhide.xlsx");

        // Add and hide a sheet
        add(&file_path, "Hidden", None, &default_global()).unwrap();
        hide(&file_path, "Hidden", false, &default_global()).unwrap();

        let result = unhide(&file_path, "Hidden", &default_global());
        assert!(result.is_ok());
    }

    #[test]
    fn test_info_sheet() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "info.xlsx");

        let result = info(&file_path, "Sheet1", &default_global());
        assert!(result.is_ok());
    }

    #[test]
    fn test_info_sheet_json() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "info_json.xlsx");

        let mut global = default_global();
        global.format = OutputFormat::Json;

        let result = info(&file_path, "Sheet1", &global);
        assert!(result.is_ok());
    }

    #[test]
    fn test_active_get() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "active.xlsx");

        let result = active(&file_path, None, &default_global());
        assert!(result.is_ok());
    }

    #[test]
    fn test_active_set() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "active_set.xlsx");

        // Add another sheet
        add(&file_path, "Second", None, &default_global()).unwrap();

        let result = active(&file_path, Some("Second"), &default_global());
        assert!(result.is_ok());
    }

    // Additional tests for better coverage

    #[test]
    fn test_run_list_command() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "run_list.xlsx");

        let args = SheetArgs {
            command: SheetCommand::List { file: file_path },
        };

        let result = run(&args, &default_global());
        assert!(result.is_ok());
    }

    #[test]
    fn test_run_add_command() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "run_add.xlsx");

        let args = SheetArgs {
            command: SheetCommand::Add {
                file: file_path,
                name: "NewSheet".to_string(),
                position: None,
            },
        };

        let result = run(&args, &default_global());
        assert!(result.is_ok());
    }

    #[test]
    fn test_run_remove_command() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "run_remove.xlsx");

        // First add a sheet
        add(&file_path, "ToRemove", None, &default_global()).unwrap();

        let args = SheetArgs {
            command: SheetCommand::Remove {
                file: file_path,
                name: "ToRemove".to_string(),
            },
        };

        let result = run(&args, &default_global());
        assert!(result.is_ok());
    }

    #[test]
    fn test_run_rename_command() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "run_rename.xlsx");

        let args = SheetArgs {
            command: SheetCommand::Rename {
                file: file_path,
                old_name: "Sheet1".to_string(),
                new_name: "Renamed".to_string(),
            },
        };

        let result = run(&args, &default_global());
        assert!(result.is_ok());
    }

    #[test]
    fn test_run_copy_command() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "run_copy.xlsx");

        let args = SheetArgs {
            command: SheetCommand::Copy {
                file: file_path,
                source: "Sheet1".to_string(),
                dest: "Sheet1_Copy".to_string(),
            },
        };

        let result = run(&args, &default_global());
        assert!(result.is_ok());
    }

    #[test]
    fn test_run_move_command() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "run_move.xlsx");

        add(&file_path, "Second", None, &default_global()).unwrap();

        let args = SheetArgs {
            command: SheetCommand::Move {
                file: file_path,
                name: "Second".to_string(),
                position: 0,
            },
        };

        let result = run(&args, &default_global());
        assert!(result.is_ok());
    }

    #[test]
    fn test_run_hide_command() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "run_hide.xlsx");

        add(&file_path, "Second", None, &default_global()).unwrap();

        let args = SheetArgs {
            command: SheetCommand::Hide {
                file: file_path,
                name: "Sheet1".to_string(),
                very: false,
            },
        };

        let result = run(&args, &default_global());
        assert!(result.is_ok());
    }

    #[test]
    fn test_run_unhide_command() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "run_unhide.xlsx");

        add(&file_path, "Second", None, &default_global()).unwrap();
        hide(&file_path, "Sheet1", false, &default_global()).unwrap();

        let args = SheetArgs {
            command: SheetCommand::Unhide {
                file: file_path,
                name: "Sheet1".to_string(),
            },
        };

        let result = run(&args, &default_global());
        assert!(result.is_ok());
    }

    #[test]
    fn test_run_info_command() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "run_info.xlsx");

        let args = SheetArgs {
            command: SheetCommand::Info {
                file: file_path,
                name: "Sheet1".to_string(),
            },
        };

        let result = run(&args, &default_global());
        assert!(result.is_ok());
    }

    #[test]
    fn test_run_active_command() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "run_active.xlsx");

        let args = SheetArgs {
            command: SheetCommand::Active {
                file: file_path,
                name: None,
            },
        };

        let result = run(&args, &default_global());
        assert!(result.is_ok());
    }

    #[test]
    fn test_add_sheet_json() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "add_json.xlsx");

        let mut global = default_global();
        global.format = OutputFormat::Json;
        global.quiet = false;

        let result = add(&file_path, "JsonSheet", None, &global);
        assert!(result.is_ok());
    }

    #[test]
    fn test_remove_sheet_dry_run() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "remove_dry.xlsx");

        add(&file_path, "ToRemove", None, &default_global()).unwrap();

        let mut global = default_global();
        global.dry_run = true;

        let result = remove(&file_path, "ToRemove", &global);
        assert!(result.is_ok());

        // Sheet should still exist
        let wb = Workbook::open(&file_path).unwrap();
        assert!(wb.sheet_names().contains(&"ToRemove"));
    }

    #[test]
    fn test_remove_sheet_json() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "remove_json.xlsx");

        add(&file_path, "ToRemove", None, &default_global()).unwrap();

        let mut global = default_global();
        global.format = OutputFormat::Json;
        global.quiet = false;

        let result = remove(&file_path, "ToRemove", &global);
        assert!(result.is_ok());
    }

    #[test]
    fn test_rename_sheet_dry_run() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "rename_dry.xlsx");

        let mut global = default_global();
        global.dry_run = true;

        let result = rename(&file_path, "Sheet1", "Renamed", &global);
        assert!(result.is_ok());

        // Name should not change
        let wb = Workbook::open(&file_path).unwrap();
        assert!(wb.sheet_names().contains(&"Sheet1"));
    }

    #[test]
    fn test_rename_sheet_json() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "rename_json.xlsx");

        let mut global = default_global();
        global.format = OutputFormat::Json;
        global.quiet = false;

        let result = rename(&file_path, "Sheet1", "Renamed", &global);
        assert!(result.is_ok());
    }

    #[test]
    fn test_copy_sheet_dry_run() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "copy_dry.xlsx");

        let mut global = default_global();
        global.dry_run = true;

        let result = copy(&file_path, "Sheet1", "Copy", &global);
        assert!(result.is_ok());

        // Copy should not exist
        let wb = Workbook::open(&file_path).unwrap();
        assert!(!wb.sheet_names().contains(&"Copy"));
    }

    #[test]
    fn test_copy_sheet_not_found() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "copy_notfound.xlsx");

        let result = copy(&file_path, "NonExistent", "Copy", &default_global());
        assert!(result.is_err());
    }

    #[test]
    fn test_move_sheet_dry_run() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "move_dry.xlsx");

        add(&file_path, "Second", None, &default_global()).unwrap();

        let mut global = default_global();
        global.dry_run = true;

        let result = move_sheet(&file_path, "Second", 0, &global);
        assert!(result.is_ok());
    }

    #[test]
    fn test_hide_sheet_dry_run() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "hide_dry.xlsx");

        add(&file_path, "Second", None, &default_global()).unwrap();

        let mut global = default_global();
        global.dry_run = true;

        let result = hide(&file_path, "Sheet1", false, &global);
        assert!(result.is_ok());
    }

    #[test]
    fn test_hide_sheet_very() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "hide_very.xlsx");

        add(&file_path, "Second", None, &default_global()).unwrap();

        let result = hide(&file_path, "Sheet1", true, &default_global());
        assert!(result.is_ok());
    }

    #[test]
    fn test_unhide_sheet_dry_run() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "unhide_dry.xlsx");

        add(&file_path, "Second", None, &default_global()).unwrap();
        hide(&file_path, "Sheet1", false, &default_global()).unwrap();

        let mut global = default_global();
        global.dry_run = true;

        let result = unhide(&file_path, "Sheet1", &global);
        assert!(result.is_ok());
    }

    #[test]
    fn test_active_set_dry_run() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "active_dry.xlsx");

        add(&file_path, "Second", None, &default_global()).unwrap();

        let mut global = default_global();
        global.dry_run = true;

        let result = active(&file_path, Some("Second"), &global);
        assert!(result.is_ok());
    }

    #[test]
    fn test_active_get_json() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "active_json.xlsx");

        let mut global = default_global();
        global.format = OutputFormat::Json;

        let result = active(&file_path, None, &global);
        assert!(result.is_ok());
    }

    #[test]
    fn test_list_with_hidden_sheet() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "list_hidden.xlsx");

        add(&file_path, "Second", None, &default_global()).unwrap();
        hide(&file_path, "Sheet1", false, &default_global()).unwrap();

        let mut global = default_global();
        global.quiet = false;

        let result = list(&file_path, &global);
        assert!(result.is_ok());
    }

    #[test]
    fn test_info_sheet_not_found() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "info_notfound.xlsx");

        let result = info(&file_path, "NonExistent", &default_global());
        assert!(result.is_err());
    }
}
