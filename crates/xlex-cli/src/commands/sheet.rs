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
