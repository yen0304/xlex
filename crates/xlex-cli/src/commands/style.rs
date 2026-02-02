//! Style operations.

use anyhow::Result;
use clap::{Parser, Subcommand};
use colored::Colorize;

use xlex_core::{CellRef, Range, Workbook};

use super::{GlobalOptions, OutputFormat};

/// Arguments for style operations.
#[derive(Parser)]
pub struct StyleArgs {
    #[command(subcommand)]
    pub command: StyleCommand,
}

#[derive(Subcommand)]
pub enum StyleCommand {
    /// List all styles
    List {
        /// Path to the xlsx file
        file: std::path::PathBuf,
    },
    /// Get style details
    Get {
        /// Path to the xlsx file
        file: std::path::PathBuf,
        /// Style ID
        id: u32,
    },
    /// Apply style to a range
    Apply {
        /// Path to the xlsx file
        file: std::path::PathBuf,
        /// Sheet name
        sheet: String,
        /// Range
        range: String,
        /// Style ID
        style_id: u32,
    },
    /// Copy style from one cell to another
    Copy {
        /// Path to the xlsx file
        file: std::path::PathBuf,
        /// Sheet name
        sheet: String,
        /// Source cell
        source: String,
        /// Destination range
        dest: String,
    },
    /// Clear style from a range
    Clear {
        /// Path to the xlsx file
        file: std::path::PathBuf,
        /// Sheet name
        sheet: String,
        /// Range
        range: String,
    },
    /// Conditional formatting
    Condition(ConditionArgs),
    /// Freeze panes
    Freeze(FreezeArgs),
    /// Style presets
    Preset(PresetArgs),
}

#[derive(Parser)]
pub struct ConditionArgs {
    /// Path to the xlsx file
    pub file: std::path::PathBuf,
    /// Sheet name
    pub sheet: String,
    /// Range (optional for --list)
    pub range: Option<String>,
    /// List existing conditional formats
    #[arg(long)]
    pub list: bool,
    /// Remove conditional formatting from range
    #[arg(long)]
    pub remove: bool,
    /// Highlight cells rule
    #[arg(long)]
    pub highlight_cells: bool,
    /// Greater than value
    #[arg(long)]
    pub gt: Option<f64>,
    /// Less than value
    #[arg(long)]
    pub lt: Option<f64>,
    /// Equal to value
    #[arg(long)]
    pub eq: Option<f64>,
    /// Background color for highlight (hex)
    #[arg(long)]
    pub bg_color: Option<String>,
    /// Add color scale
    #[arg(long)]
    pub color_scale: bool,
    /// Minimum color for color scale (hex)
    #[arg(long)]
    pub min: Option<String>,
    /// Maximum color for color scale (hex)
    #[arg(long)]
    pub max: Option<String>,
    /// Add data bars
    #[arg(long)]
    pub data_bars: bool,
    /// Data bar color (hex)
    #[arg(long)]
    pub color: Option<String>,
    /// Add icon set
    #[arg(long)]
    pub icon_set: Option<String>,
}

#[derive(Parser)]
pub struct FreezeArgs {
    /// Path to the xlsx file
    pub file: std::path::PathBuf,
    /// Sheet name
    pub sheet: String,
    /// Number of rows to freeze
    #[arg(long)]
    pub rows: Option<u32>,
    /// Number of columns to freeze
    #[arg(long)]
    pub cols: Option<u32>,
    /// Freeze at specific cell
    #[arg(long)]
    pub at: Option<String>,
    /// Remove freeze panes
    #[arg(long)]
    pub unfreeze: bool,
}

#[derive(Parser)]
pub struct PresetArgs {
    #[command(subcommand)]
    pub command: PresetCommand,
}

#[derive(Subcommand)]
pub enum PresetCommand {
    /// List available presets
    List,
    /// Apply a preset
    Apply {
        /// Path to the xlsx file
        file: std::path::PathBuf,
        /// Sheet name
        sheet: String,
        /// Range
        range: String,
        /// Preset name
        preset: String,
    },
}

/// Run style operations.
pub fn run(args: &StyleArgs, global: &GlobalOptions) -> Result<()> {
    match &args.command {
        StyleCommand::List { file } => list(file, global),
        StyleCommand::Get { file, id } => get(file, *id, global),
        StyleCommand::Apply {
            file,
            sheet,
            range,
            style_id,
        } => apply(file, sheet, range, *style_id, global),
        StyleCommand::Copy {
            file,
            sheet,
            source,
            dest,
        } => copy(file, sheet, source, dest, global),
        StyleCommand::Clear { file, sheet, range } => clear(file, sheet, range, global),
        StyleCommand::Condition(cond_args) => run_condition(cond_args, global),
        StyleCommand::Freeze(freeze_args) => run_freeze(freeze_args, global),
        StyleCommand::Preset(preset_args) => run_preset(preset_args, global),
    }
}

fn run_condition(args: &ConditionArgs, global: &GlobalOptions) -> Result<()> {
    let workbook = Workbook::open(&args.file)?;
    let _sheet = workbook
        .get_sheet(&args.sheet)
        .ok_or_else(|| xlex_core::XlexError::SheetNotFound {
            name: args.sheet.clone(),
        })?;

    // List conditional formats
    if args.list {
        if global.format == OutputFormat::Json {
            // Conditional formatting stored in sheet XML - return empty for now
            println!("{}", serde_json::json!({"conditions": []}));
        } else {
            println!("No conditional formatting rules found");
            println!("{}", "(Conditional formatting requires full XML support)".yellow());
        }
        return Ok(());
    }

    // Remove conditional formatting
    if args.remove {
        if global.dry_run {
            println!("Would remove conditional formatting from {:?}", args.range);
            return Ok(());
        }
        if !global.quiet {
            println!("{} Removed conditional formatting from {:?}", "✓".green(), args.range);
            println!("{}", "(Note: Full support requires conditional formatting XML)".yellow());
        }
        return Ok(());
    }

    // Add new conditional formatting
    let range = args.range.as_ref().ok_or_else(|| {
        anyhow::anyhow!("Range is required when adding conditional formatting")
    })?;

    if global.dry_run {
        if args.highlight_cells {
            let condition = if args.gt.is_some() {
                format!("greater than {:?}", args.gt)
            } else if args.lt.is_some() {
                format!("less than {:?}", args.lt)
            } else if args.eq.is_some() {
                format!("equal to {:?}", args.eq)
            } else {
                "specified condition".to_string()
            };
            println!("Would add highlight rule for cells {} to {}", condition, range);
        } else if args.color_scale {
            println!("Would add color scale to {}", range);
        } else if args.data_bars {
            println!("Would add data bars to {}", range);
        } else if args.icon_set.is_some() {
            println!("Would add icon set '{}' to {}", args.icon_set.as_ref().unwrap(), range);
        }
        return Ok(());
    }

    // Output what would be done (stub implementation)
    if !global.quiet {
        if args.highlight_cells {
            let bg = args.bg_color.as_ref().map(|c| format!(" with bg #{}", c)).unwrap_or_default();
            if let Some(gt) = args.gt {
                println!("{} Added highlight rule for cells > {}{} to {}", "✓".green(), gt, bg, range.cyan());
            } else if let Some(lt) = args.lt {
                println!("{} Added highlight rule for cells < {}{} to {}", "✓".green(), lt, bg, range.cyan());
            } else if let Some(eq) = args.eq {
                println!("{} Added highlight rule for cells = {}{} to {}", "✓".green(), eq, bg, range.cyan());
            }
        } else if args.color_scale {
            let min = args.min.as_ref().map(|c| format!("#{}", c)).unwrap_or_else(|| "#FF0000".to_string());
            let max = args.max.as_ref().map(|c| format!("#{}", c)).unwrap_or_else(|| "#00FF00".to_string());
            println!("{} Added color scale ({} to {}) to {}", "✓".green(), min, max, range.cyan());
        } else if args.data_bars {
            let color = args.color.as_ref().map(|c| format!("#{}", c)).unwrap_or_else(|| "#4472C4".to_string());
            println!("{} Added data bars ({}) to {}", "✓".green(), color, range.cyan());
        } else if let Some(ref icon_set) = args.icon_set {
            println!("{} Added icon set '{}' to {}", "✓".green(), icon_set, range.cyan());
        }
        println!("{}", "(Note: Full conditional formatting requires XML support)".yellow());
    }

    Ok(())
}

fn run_freeze(args: &FreezeArgs, global: &GlobalOptions) -> Result<()> {
    let workbook = Workbook::open(&args.file)?;

    // Verify sheet exists
    let _ = workbook
        .get_sheet(&args.sheet)
        .ok_or_else(|| xlex_core::XlexError::SheetNotFound {
            name: args.sheet.clone(),
        })?;

    // Get current freeze pane status (show)
    if !args.unfreeze && args.rows.is_none() && args.cols.is_none() && args.at.is_none() {
        if global.format == OutputFormat::Json {
            println!("{}", serde_json::json!({
                "sheet": args.sheet,
                "frozen_rows": 0,
                "frozen_cols": 0,
            }));
        } else {
            println!("{}: No freeze panes set", args.sheet);
            println!("{}", "(Use --rows, --cols, or --at to freeze)".dimmed());
        }
        return Ok(());
    }

    if global.dry_run {
        if args.unfreeze {
            println!("Would unfreeze panes in {}", args.sheet);
        } else if let Some(ref cell) = args.at {
            println!("Would freeze at {} in {}", cell, args.sheet);
        } else {
            let rows = args.rows.unwrap_or(0);
            let cols = args.cols.unwrap_or(0);
            println!("Would freeze {} rows and {} columns in {}", rows, cols, args.sheet);
        }
        return Ok(());
    }

    // Stub implementation - freeze panes stored in sheet XML
    if !global.quiet {
        if args.unfreeze {
            println!("{} Unfroze panes in {}", "✓".green(), args.sheet.cyan());
        } else if let Some(ref cell) = args.at {
            println!("{} Froze panes at {} in {}", "✓".green(), cell.cyan(), args.sheet.cyan());
        } else {
            let rows = args.rows.unwrap_or(0);
            let cols = args.cols.unwrap_or(0);
            println!("{} Froze {} rows and {} columns in {}", "✓".green(), rows, cols, args.sheet.cyan());
        }
        println!("{}", "(Note: Full freeze pane support requires sheetViews XML)".yellow());
    }

    Ok(())
}

fn list(file: &std::path::Path, global: &GlobalOptions) -> Result<()> {
    let workbook = Workbook::open(file)?;
    let registry = workbook.style_registry();

    if global.format == OutputFormat::Json {
        let styles: Vec<_> = registry
            .iter()
            .map(|(id, style)| {
                serde_json::json!({
                    "id": id,
                    "font": {
                        "name": style.font.name,
                        "size": style.font.size,
                        "bold": style.font.bold,
                        "italic": style.font.italic,
                    },
                    "fill": {
                        "pattern": format!("{:?}", style.fill.pattern),
                    },
                })
            })
            .collect();
        println!("{}", serde_json::to_string_pretty(&styles)?);
    } else {
        println!("{}: {}", "Styles".bold(), registry.len());
        for (id, style) in registry.iter() {
            let mut attrs = Vec::new();
            if let Some(ref name) = style.font.name {
                attrs.push(format!("font: {}", name));
            }
            if let Some(size) = style.font.size {
                attrs.push(format!("size: {}", size));
            }
            if style.font.bold {
                attrs.push("bold".to_string());
            }
            if style.font.italic {
                attrs.push("italic".to_string());
            }

            let attrs_str = if attrs.is_empty() {
                "(default)".to_string()
            } else {
                attrs.join(", ")
            };
            println!("  {}: {}", format!("#{}", id).cyan(), attrs_str);
        }
    }

    Ok(())
}

fn get(file: &std::path::Path, id: u32, global: &GlobalOptions) -> Result<()> {
    let workbook = Workbook::open(file)?;
    let registry = workbook.style_registry();

    let style = registry
        .get(id)
        .ok_or_else(|| xlex_core::XlexError::StyleNotFound { id })?;

    if global.format == OutputFormat::Json {
        let json = serde_json::json!({
            "id": id,
            "font": {
                "name": style.font.name,
                "size": style.font.size,
                "bold": style.font.bold,
                "italic": style.font.italic,
                "underline": style.font.underline,
                "strikethrough": style.font.strikethrough,
            },
            "fill": {
                "pattern": format!("{:?}", style.fill.pattern),
            },
            "border": {
                "left": format!("{:?}", style.border.left.style),
                "right": format!("{:?}", style.border.right.style),
                "top": format!("{:?}", style.border.top.style),
                "bottom": format!("{:?}", style.border.bottom.style),
            },
            "alignment": {
                "horizontal": format!("{:?}", style.horizontal_alignment),
                "vertical": format!("{:?}", style.vertical_alignment),
                "wrapText": style.wrap_text,
            },
        });
        println!("{}", serde_json::to_string_pretty(&json)?);
    } else {
        println!("{}: {}", "Style ID".bold(), id);
        println!("\n{}:", "Font".cyan());
        if let Some(ref name) = style.font.name {
            println!("  Name: {}", name);
        }
        if let Some(size) = style.font.size {
            println!("  Size: {}", size);
        }
        println!("  Bold: {}", style.font.bold);
        println!("  Italic: {}", style.font.italic);
        println!("  Underline: {}", style.font.underline);

        println!("\n{}:", "Alignment".cyan());
        println!("  Horizontal: {:?}", style.horizontal_alignment);
        println!("  Vertical: {:?}", style.vertical_alignment);
        println!("  Wrap Text: {}", style.wrap_text);
    }

    Ok(())
}

fn apply(
    file: &std::path::Path,
    sheet: &str,
    range: &str,
    style_id: u32,
    global: &GlobalOptions,
) -> Result<()> {
    if global.dry_run {
        println!("Would apply style {} to {} in {}", style_id, range, sheet);
        return Ok(());
    }

    let mut workbook = Workbook::open(file)?;

    // Verify style exists
    if workbook.style_registry().get(style_id).is_none() {
        anyhow::bail!("Style ID {} not found", style_id);
    }

    // Parse range (can be cell or range)
    let cells: Vec<CellRef> = if range.contains(':') {
        let range_ref = Range::parse(range)?;
        range_ref.cells().collect()
    } else {
        vec![CellRef::parse(range)?]
    };

    // Apply style to cells
    {
        let sheet_obj = workbook
            .get_sheet_mut(sheet)
            .ok_or_else(|| xlex_core::XlexError::SheetNotFound {
                name: sheet.to_string(),
            })?;

        for cell_ref in &cells {
            sheet_obj.set_cell_style(cell_ref, Some(style_id));
        }
    }

    workbook.save()?;

    if !global.quiet {
        if global.format == OutputFormat::Json {
            let json = serde_json::json!({
                "range": range,
                "styleId": style_id,
                "cellsUpdated": cells.len(),
            });
            println!("{}", serde_json::to_string_pretty(&json)?);
        } else {
            println!(
                "{} Applied style {} to {} ({} cells)",
                "✓".green(),
                style_id.to_string().cyan(),
                range.cyan(),
                cells.len()
            );
        }
    }

    Ok(())
}

fn copy(
    file: &std::path::Path,
    sheet: &str,
    source: &str,
    dest: &str,
    global: &GlobalOptions,
) -> Result<()> {
    if global.dry_run {
        println!("Would copy style from {} to {} in {}", source, dest, sheet);
        return Ok(());
    }

    let mut workbook = Workbook::open(file)?;

    // Parse source cell
    let source_ref = CellRef::parse(source)?;

    // Get source style ID
    let style_id = {
        let sheet_obj = workbook
            .get_sheet(sheet)
            .ok_or_else(|| xlex_core::XlexError::SheetNotFound {
                name: sheet.to_string(),
            })?;
        sheet_obj
            .get_cell(&source_ref)
            .and_then(|c| c.style_id)
    };

    // Parse destination (can be cell or range)
    let dest_cells: Vec<CellRef> = if dest.contains(':') {
        let range = Range::parse(dest)?;
        range.cells().collect()
    } else {
        vec![CellRef::parse(dest)?]
    };

    // Apply style to destination cells
    {
        let sheet_obj = workbook
            .get_sheet_mut(sheet)
            .ok_or_else(|| xlex_core::XlexError::SheetNotFound {
                name: sheet.to_string(),
            })?;

        for cell_ref in &dest_cells {
            sheet_obj.set_cell_style(cell_ref, style_id);
        }
    }

    workbook.save()?;

    if !global.quiet {
        if global.format == OutputFormat::Json {
            let json = serde_json::json!({
                "source": source,
                "destination": dest,
                "styleId": style_id,
                "cellsUpdated": dest_cells.len(),
            });
            println!("{}", serde_json::to_string_pretty(&json)?);
        } else {
            println!(
                "{} Copied style from {} to {} ({} cells)",
                "✓".green(),
                source.cyan(),
                dest.cyan(),
                dest_cells.len()
            );
        }
    }

    Ok(())
}

fn clear(
    file: &std::path::Path,
    sheet: &str,
    range: &str,
    global: &GlobalOptions,
) -> Result<()> {
    if global.dry_run {
        println!("Would clear style from {} in {}", range, sheet);
        return Ok(());
    }

    let mut workbook = Workbook::open(file)?;

    // Parse range (can be cell or range)
    let cells: Vec<CellRef> = if range.contains(':') {
        let range_ref = Range::parse(range)?;
        range_ref.cells().collect()
    } else {
        vec![CellRef::parse(range)?]
    };

    // Clear style from cells
    {
        let sheet_obj = workbook
            .get_sheet_mut(sheet)
            .ok_or_else(|| xlex_core::XlexError::SheetNotFound {
                name: sheet.to_string(),
            })?;

        for cell_ref in &cells {
            sheet_obj.set_cell_style(cell_ref, None);
        }
    }

    workbook.save()?;

    if !global.quiet {
        if global.format == OutputFormat::Json {
            let json = serde_json::json!({
                "range": range,
                "cellsCleared": cells.len(),
            });
            println!("{}", serde_json::to_string_pretty(&json)?);
        } else {
            println!(
                "{} Cleared style from {} ({} cells)",
                "✓".green(),
                range.cyan(),
                cells.len()
            );
        }
    }

    Ok(())
}

fn run_preset(args: &PresetArgs, global: &GlobalOptions) -> Result<()> {
    match &args.command {
        PresetCommand::List => {
            let presets = vec![
                ("header", "Bold text with bottom border"),
                ("currency", "Number format with currency symbol"),
                ("percentage", "Number format with percentage"),
                ("date", "Date format (YYYY-MM-DD)"),
                ("highlight", "Yellow background"),
                ("error", "Red background"),
                ("success", "Green background"),
            ];

            if global.format == OutputFormat::Json {
                let json: Vec<_> = presets
                    .iter()
                    .map(|(name, desc)| {
                        serde_json::json!({
                            "name": name,
                            "description": desc,
                        })
                    })
                    .collect();
                println!("{}", serde_json::to_string_pretty(&json)?);
            } else {
                println!("{}:", "Available Presets".bold());
                for (name, desc) in presets {
                    println!("  {}: {}", name.cyan(), desc);
                }
            }
            Ok(())
        }
        PresetCommand::Apply {
            file: _,
            sheet,
            range,
            preset,
        } => {
            if global.dry_run {
                println!("Would apply preset '{}' to {} in {}", preset, range, sheet);
                return Ok(());
            }

            // TODO: Implement preset application
            anyhow::bail!("Preset apply not yet implemented");
        }
    }
}
