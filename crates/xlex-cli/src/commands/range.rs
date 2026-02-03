//! Range operations.

use anyhow::Result;
use clap::{Parser, Subcommand};
use colored::Colorize;

use xlex_core::{DefinedName, Range, Workbook};

use super::{GlobalOptions, OutputFormat};

/// Arguments for range operations.
#[derive(Parser)]
pub struct RangeArgs {
    #[command(subcommand)]
    pub command: RangeCommand,
}

#[derive(Subcommand)]
pub enum RangeCommand {
    /// Get range data
    Get {
        /// Path to the xlsx file
        file: std::path::PathBuf,
        /// Sheet name
        sheet: String,
        /// Range (e.g., A1:B10)
        range: String,
    },
    /// Copy a range
    Copy {
        /// Path to the xlsx file
        file: std::path::PathBuf,
        /// Sheet name
        sheet: String,
        /// Source range
        source: String,
        /// Destination (top-left cell)
        dest: String,
    },
    /// Move a range
    Move {
        /// Path to the xlsx file
        file: std::path::PathBuf,
        /// Sheet name
        sheet: String,
        /// Source range
        source: String,
        /// Destination (top-left cell)
        dest: String,
    },
    /// Clear a range
    Clear {
        /// Path to the xlsx file
        file: std::path::PathBuf,
        /// Sheet name
        sheet: String,
        /// Range
        range: String,
        /// Clear values only (keep formatting)
        #[arg(long)]
        values_only: bool,
    },
    /// Fill a range with a pattern
    Fill {
        /// Path to the xlsx file
        file: std::path::PathBuf,
        /// Sheet name
        sheet: String,
        /// Range
        range: String,
        /// Value or pattern
        value: String,
    },
    /// Merge cells in a range
    Merge {
        /// Path to the xlsx file
        file: std::path::PathBuf,
        /// Sheet name
        sheet: String,
        /// Range
        range: String,
    },
    /// Unmerge cells in a range
    Unmerge {
        /// Path to the xlsx file
        file: std::path::PathBuf,
        /// Sheet name
        sheet: String,
        /// Range
        range: String,
    },
    /// Apply styling to a range
    Style {
        /// Path to the xlsx file
        file: std::path::PathBuf,
        /// Sheet name
        sheet: String,
        /// Range
        range: String,
        /// Bold text
        #[arg(long)]
        bold: bool,
        /// Italic text
        #[arg(long)]
        italic: bool,
        /// Underline text
        #[arg(long)]
        underline: bool,
        /// Font name
        #[arg(long)]
        font: Option<String>,
        /// Font size
        #[arg(long)]
        font_size: Option<f64>,
        /// Text color (hex, e.g., FF0000)
        #[arg(long)]
        text_color: Option<String>,
        /// Background color (hex, e.g., FFFF00)
        #[arg(long)]
        bg_color: Option<String>,
        /// Horizontal alignment (left, center, right)
        #[arg(long)]
        align: Option<String>,
        /// Vertical alignment (top, middle, bottom)
        #[arg(long)]
        valign: Option<String>,
        /// Enable text wrap
        #[arg(long)]
        wrap: bool,
        /// Number format (e.g., #,##0.00)
        #[arg(long)]
        number_format: Option<String>,
        /// Format as percentage
        #[arg(long)]
        percent: bool,
        /// Format as currency
        #[arg(long)]
        currency: Option<String>,
        /// Date format (e.g., YYYY-MM-DD)
        #[arg(long)]
        date_format: Option<String>,
    },
    /// Apply borders to a range
    Border {
        /// Path to the xlsx file
        file: std::path::PathBuf,
        /// Sheet name
        sheet: String,
        /// Range
        range: String,
        /// Apply borders to all cells
        #[arg(long)]
        all: bool,
        /// Apply outline border only
        #[arg(long)]
        outline: bool,
        /// Apply top border
        #[arg(long)]
        top: bool,
        /// Apply bottom border
        #[arg(long)]
        bottom: bool,
        /// Apply left border
        #[arg(long)]
        left: bool,
        /// Apply right border
        #[arg(long)]
        right: bool,
        /// Remove all borders
        #[arg(long)]
        none: bool,
        /// Border style (thin, medium, thick, dashed, dotted, double)
        #[arg(long, default_value = "thin")]
        style: String,
        /// Border color (hex, e.g., 000000)
        #[arg(long)]
        border_color: Option<String>,
    },
    /// Define a named range
    Name {
        /// Path to the xlsx file
        file: std::path::PathBuf,
        /// Name for the range
        name: String,
        /// Range reference
        range: String,
        /// Sheet scope (global if not specified)
        #[arg(long)]
        sheet: Option<String>,
    },
    /// List named ranges
    Names {
        /// Path to the xlsx file
        file: std::path::PathBuf,
    },
    /// Validate range data
    Validate {
        /// Path to the xlsx file
        file: std::path::PathBuf,
        /// Sheet name
        sheet: String,
        /// Range
        range: String,
        /// Validation rule (nonempty, numeric, unique)
        rule: String,
    },
    /// Sort range
    Sort {
        /// Path to the xlsx file
        file: std::path::PathBuf,
        /// Sheet name
        sheet: String,
        /// Range
        range: String,
        /// Column to sort by (within range)
        #[arg(long, short = 'c')]
        column: Option<String>,
        /// Sort descending
        #[arg(long, short = 'd')]
        descending: bool,
    },
    /// Filter range
    Filter {
        /// Path to the xlsx file
        file: std::path::PathBuf,
        /// Sheet name
        sheet: String,
        /// Range
        range: String,
        /// Column to filter
        column: String,
        /// Filter value
        value: String,
    },
}

/// Run range operations.
pub fn run(args: &RangeArgs, global: &GlobalOptions) -> Result<()> {
    match &args.command {
        RangeCommand::Get { file, sheet, range } => get(file, sheet, range, global),
        RangeCommand::Copy {
            file,
            sheet,
            source,
            dest,
        } => copy(file, sheet, source, dest, global),
        RangeCommand::Move {
            file,
            sheet,
            source,
            dest,
        } => move_range(file, sheet, source, dest, global),
        RangeCommand::Clear {
            file,
            sheet,
            range,
            values_only,
        } => clear(file, sheet, range, *values_only, global),
        RangeCommand::Fill {
            file,
            sheet,
            range,
            value,
        } => fill(file, sheet, range, value, global),
        RangeCommand::Merge { file, sheet, range } => merge(file, sheet, range, global),
        RangeCommand::Unmerge { file, sheet, range } => unmerge(file, sheet, range, global),
        RangeCommand::Style {
            file,
            sheet,
            range,
            bold,
            italic,
            underline,
            font,
            font_size,
            text_color,
            bg_color,
            align,
            valign,
            wrap,
            number_format,
            percent,
            currency,
            date_format,
        } => range_style(
            file,
            sheet,
            range,
            RangeStyleOpts {
                bold: *bold,
                italic: *italic,
                underline: *underline,
                font: font.clone(),
                font_size: *font_size,
                color: text_color.clone(),
                bg_color: bg_color.clone(),
                align: align.clone(),
                valign: valign.clone(),
                wrap: *wrap,
                number_format: number_format.clone(),
                percent: *percent,
                currency: currency.clone(),
                date_format: date_format.clone(),
            },
            global,
        ),
        RangeCommand::Border {
            file,
            sheet,
            range,
            all,
            outline,
            top,
            bottom,
            left,
            right,
            none,
            style,
            border_color,
        } => range_border(
            file,
            sheet,
            range,
            RangeBorderOpts {
                all: *all,
                outline: *outline,
                top: *top,
                bottom: *bottom,
                left: *left,
                right: *right,
                none: *none,
                style: style.clone(),
                border_color: border_color.clone(),
            },
            global,
        ),
        RangeCommand::Name {
            file,
            name: range_name,
            range,
            sheet,
        } => name(file, range_name, range, sheet.as_deref(), global),
        RangeCommand::Names { file } => names(file, global),
        RangeCommand::Validate {
            file,
            sheet,
            range,
            rule,
        } => validate(file, sheet, range, rule, global),
        RangeCommand::Sort {
            file,
            sheet,
            range,
            column,
            descending,
        } => sort(file, sheet, range, column.as_deref(), *descending, global),
        RangeCommand::Filter {
            file,
            sheet,
            range,
            column,
            value,
        } => filter(file, sheet, range, column, value, global),
    }
}

#[derive(Default)]
struct RangeStyleOpts {
    bold: bool,
    italic: bool,
    underline: bool,
    font: Option<String>,
    font_size: Option<f64>,
    color: Option<String>,
    bg_color: Option<String>,
    align: Option<String>,
    valign: Option<String>,
    wrap: bool,
    number_format: Option<String>,
    percent: bool,
    currency: Option<String>,
    date_format: Option<String>,
}

#[derive(Default)]
struct RangeBorderOpts {
    all: bool,
    outline: bool,
    top: bool,
    bottom: bool,
    left: bool,
    right: bool,
    none: bool,
    style: String,
    border_color: Option<String>,
}

fn range_style(
    file: &std::path::Path,
    sheet: &str,
    range: &str,
    opts: RangeStyleOpts,
    global: &GlobalOptions,
) -> Result<()> {
    use xlex_core::{
        style::{Color, FillPattern, HorizontalAlignment, NumberFormat, Style, VerticalAlignment},
        CellRef,
    };

    if global.dry_run {
        println!("Would apply styles to range {}", range);
        return Ok(());
    }

    let mut workbook = Workbook::open(file)?;
    let range_ref = Range::parse(range)?;

    // Build style from options
    let mut style = Style::default();

    // Font settings
    style.font.bold = opts.bold;
    style.font.italic = opts.italic;
    style.font.underline = opts.underline;
    style.font.name = opts.font;
    style.font.size = opts.font_size;

    if let Some(ref color_str) = opts.color {
        style.font.color = Color::from_hex(color_str);
    }

    // Fill settings
    if let Some(ref bg_color_str) = opts.bg_color {
        if let Some(color) = Color::from_hex(bg_color_str) {
            style.fill.pattern = FillPattern::Solid;
            style.fill.fg_color = Some(color);
        }
    }

    // Alignment
    if let Some(ref align) = opts.align {
        style.horizontal_alignment = match align.to_lowercase().as_str() {
            "left" => HorizontalAlignment::Left,
            "center" => HorizontalAlignment::Center,
            "right" => HorizontalAlignment::Right,
            "justify" => HorizontalAlignment::Justify,
            _ => HorizontalAlignment::General,
        };
    }

    if let Some(ref valign) = opts.valign {
        style.vertical_alignment = match valign.to_lowercase().as_str() {
            "top" => VerticalAlignment::Top,
            "middle" | "center" => VerticalAlignment::Center,
            "bottom" => VerticalAlignment::Bottom,
            _ => VerticalAlignment::Center,
        };
    }

    style.wrap_text = opts.wrap;

    // Number format
    if opts.percent {
        style.number_format = NumberFormat::percentage(2);
    } else if let Some(ref _currency) = opts.currency {
        style.number_format = NumberFormat::custom("$#,##0.00");
    } else if let Some(ref _date_fmt) = opts.date_format {
        style.number_format = NumberFormat::date();
    } else if let Some(ref fmt) = opts.number_format {
        style.number_format = NumberFormat::custom(fmt.clone());
    }

    // Register the style and get its ID
    let style_id = workbook.style_registry_mut().add(style);

    {
        let sheet_obj =
            workbook
                .get_sheet_mut(sheet)
                .ok_or_else(|| xlex_core::XlexError::SheetNotFound {
                    name: sheet.to_string(),
                })?;

        // Apply style to each cell in range
        for row in range_ref.start.row..=range_ref.end.row {
            for col in range_ref.start.col..=range_ref.end.col {
                let cell_ref = CellRef::new(col, row);
                sheet_obj.set_cell_style(&cell_ref, Some(style_id));
            }
        }
    }

    workbook.save()?;

    if !global.quiet {
        println!("{} Applied styles to range {}", "✓".green(), range.cyan());
    }
    Ok(())
}

fn range_border(
    file: &std::path::Path,
    sheet: &str,
    range: &str,
    opts: RangeBorderOpts,
    global: &GlobalOptions,
) -> Result<()> {
    use xlex_core::{
        style::{Border, BorderSide, BorderStyle, Color, Style},
        CellRef,
    };

    if global.dry_run {
        println!("Would apply borders to range {}", range);
        return Ok(());
    }

    let mut workbook = Workbook::open(file)?;
    let range_ref = Range::parse(range)?;

    // Parse border style
    let border_style = match opts.style.to_lowercase().as_str() {
        "thin" => BorderStyle::Thin,
        "medium" => BorderStyle::Medium,
        "thick" => BorderStyle::Thick,
        "dashed" => BorderStyle::Dashed,
        "dotted" => BorderStyle::Dotted,
        "double" => BorderStyle::Double,
        "hair" => BorderStyle::Hair,
        _ => BorderStyle::Thin,
    };

    let border_color = opts.border_color.as_ref().and_then(|c| Color::from_hex(c));

    let border_side = BorderSide {
        style: border_style,
        color: border_color.clone(),
    };

    // Verify sheet exists
    let _ = workbook
        .get_sheet(sheet)
        .ok_or_else(|| xlex_core::XlexError::SheetNotFound {
            name: sheet.to_string(),
        })?;

    // For each cell, create a style with the appropriate border and register it
    for row in range_ref.start.row..=range_ref.end.row {
        for col in range_ref.start.col..=range_ref.end.col {
            let cell_ref = CellRef::new(col, row);

            let is_top_edge = row == range_ref.start.row;
            let is_bottom_edge = row == range_ref.end.row;
            let is_left_edge = col == range_ref.start.col;
            let is_right_edge = col == range_ref.end.col;

            let mut cell_border = Border::default();

            if opts.none {
                // Remove all borders - use default (no borders)
            } else if opts.all {
                cell_border = Border::all(border_style, border_color.clone());
            } else if opts.outline {
                if is_top_edge {
                    cell_border.top = border_side.clone();
                }
                if is_bottom_edge {
                    cell_border.bottom = border_side.clone();
                }
                if is_left_edge {
                    cell_border.left = border_side.clone();
                }
                if is_right_edge {
                    cell_border.right = border_side.clone();
                }
            } else {
                // Individual borders
                if opts.top && is_top_edge {
                    cell_border.top = border_side.clone();
                }
                if opts.bottom && is_bottom_edge {
                    cell_border.bottom = border_side.clone();
                }
                if opts.left && is_left_edge {
                    cell_border.left = border_side.clone();
                }
                if opts.right && is_right_edge {
                    cell_border.right = border_side.clone();
                }
            }

            // Create style with border
            let mut style = Style::default();
            style.border = cell_border;

            // Register style and apply to cell
            let style_id = workbook.style_registry_mut().add(style);

            let sheet_obj = workbook.get_sheet_mut(sheet).ok_or_else(|| {
                xlex_core::XlexError::SheetNotFound {
                    name: sheet.to_string(),
                }
            })?;
            sheet_obj.set_cell_style(&cell_ref, Some(style_id));
        }
    }

    workbook.save()?;

    if !global.quiet {
        let action = if opts.none { "Removed" } else { "Applied" };
        println!(
            "{} {} borders on range {}",
            "✓".green(),
            action,
            range.cyan()
        );
    }
    Ok(())
}

fn get(file: &std::path::Path, sheet: &str, range: &str, global: &GlobalOptions) -> Result<()> {
    let workbook = Workbook::open(file)?;
    let range_ref = Range::parse(range)?;
    let sheet_obj =
        workbook
            .get_sheet(sheet)
            .ok_or_else(|| xlex_core::XlexError::SheetNotFound {
                name: sheet.to_string(),
            })?;

    let mut rows: Vec<Vec<serde_json::Value>> = Vec::new();

    for row in range_ref.start.row..=range_ref.end.row {
        let mut row_values: Vec<serde_json::Value> = Vec::new();
        for col in range_ref.start.col..=range_ref.end.col {
            let cell_ref = xlex_core::CellRef::new(col, row);
            let value = sheet_obj.get_value(&cell_ref);
            row_values.push(match value {
                xlex_core::CellValue::Empty => serde_json::Value::Null,
                xlex_core::CellValue::String(s) => serde_json::Value::String(s),
                xlex_core::CellValue::Number(n) => serde_json::json!(n),
                xlex_core::CellValue::Boolean(b) => serde_json::Value::Bool(b),
                _ => serde_json::Value::String(value.to_display_string()),
            });
        }
        rows.push(row_values);
    }

    if global.format == OutputFormat::Json {
        let json = serde_json::json!({
            "range": range,
            "data": rows,
        });
        println!("{}", serde_json::to_string_pretty(&json)?);
    } else if global.format == OutputFormat::Csv {
        for row in rows {
            let values: Vec<String> = row
                .iter()
                .map(|v| match v {
                    serde_json::Value::String(s) => s.clone(),
                    serde_json::Value::Number(n) => n.to_string(),
                    serde_json::Value::Bool(b) => b.to_string(),
                    serde_json::Value::Null => String::new(),
                    _ => v.to_string(),
                })
                .collect();
            println!("{}", values.join(","));
        }
    } else {
        for (i, row) in rows.iter().enumerate() {
            let row_num = range_ref.start.row + i as u32;
            let values: Vec<String> = row
                .iter()
                .map(|v| match v {
                    serde_json::Value::String(s) => s.clone(),
                    serde_json::Value::Number(n) => n.to_string(),
                    serde_json::Value::Bool(b) => b.to_string(),
                    serde_json::Value::Null => String::new(),
                    _ => v.to_string(),
                })
                .collect();
            println!(
                "{}: {}",
                format!("Row {}", row_num).cyan(),
                values.join(" | ")
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
        println!("Would copy {} to {} in {}", source, dest, sheet);
        return Ok(());
    }

    let mut workbook = Workbook::open(file)?;
    let source_range = Range::parse(source)?;
    let dest_cell = xlex_core::CellRef::parse(dest)?;

    let sheet_obj =
        workbook
            .get_sheet(sheet)
            .ok_or_else(|| xlex_core::XlexError::SheetNotFound {
                name: sheet.to_string(),
            })?;

    // Collect source values
    let mut values: Vec<(u32, u32, xlex_core::CellValue)> = Vec::new();
    for row_offset in 0..=(source_range.end.row - source_range.start.row) {
        for col_offset in 0..=(source_range.end.col - source_range.start.col) {
            let src_ref = xlex_core::CellRef::new(
                source_range.start.col + col_offset,
                source_range.start.row + row_offset,
            );
            let value = sheet_obj.get_value(&src_ref);
            if !value.is_empty() {
                values.push((col_offset, row_offset, value));
            }
        }
    }
    let _ = sheet_obj;

    // Paste values to destination
    let sheet_obj =
        workbook
            .get_sheet_mut(sheet)
            .ok_or_else(|| xlex_core::XlexError::SheetNotFound {
                name: sheet.to_string(),
            })?;

    for (col_offset, row_offset, value) in values {
        let dest_ref =
            xlex_core::CellRef::new(dest_cell.col + col_offset, dest_cell.row + row_offset);
        sheet_obj.set_cell(dest_ref, value);
    }

    let _ = sheet_obj;
    workbook.save()?;

    if !global.quiet {
        println!("Copied {} to {}", source.cyan(), dest.green());
    }

    Ok(())
}

fn move_range(
    file: &std::path::Path,
    sheet: &str,
    source: &str,
    dest: &str,
    global: &GlobalOptions,
) -> Result<()> {
    if global.dry_run {
        println!("Would move {} to {} in {}", source, dest, sheet);
        return Ok(());
    }

    let mut workbook = Workbook::open(file)?;
    let source_range = Range::parse(source)?;
    let dest_cell = xlex_core::CellRef::parse(dest)?;

    let sheet_obj =
        workbook
            .get_sheet(sheet)
            .ok_or_else(|| xlex_core::XlexError::SheetNotFound {
                name: sheet.to_string(),
            })?;

    // Collect source values
    let mut values: Vec<(u32, u32, xlex_core::CellValue)> = Vec::new();
    for row_offset in 0..=(source_range.end.row - source_range.start.row) {
        for col_offset in 0..=(source_range.end.col - source_range.start.col) {
            let src_ref = xlex_core::CellRef::new(
                source_range.start.col + col_offset,
                source_range.start.row + row_offset,
            );
            let value = sheet_obj.get_value(&src_ref);
            values.push((col_offset, row_offset, value));
        }
    }
    let _ = sheet_obj;

    let sheet_obj =
        workbook
            .get_sheet_mut(sheet)
            .ok_or_else(|| xlex_core::XlexError::SheetNotFound {
                name: sheet.to_string(),
            })?;

    // Clear source range
    for row_offset in 0..=(source_range.end.row - source_range.start.row) {
        for col_offset in 0..=(source_range.end.col - source_range.start.col) {
            let src_ref = xlex_core::CellRef::new(
                source_range.start.col + col_offset,
                source_range.start.row + row_offset,
            );
            sheet_obj.clear_cell(&src_ref);
        }
    }

    // Paste values to destination
    for (col_offset, row_offset, value) in values {
        if !value.is_empty() {
            let dest_ref =
                xlex_core::CellRef::new(dest_cell.col + col_offset, dest_cell.row + row_offset);
            sheet_obj.set_cell(dest_ref, value);
        }
    }

    let _ = sheet_obj;
    workbook.save()?;

    if !global.quiet {
        println!("Moved {} to {}", source.cyan(), dest.green());
    }

    Ok(())
}

fn clear(
    file: &std::path::Path,
    sheet: &str,
    range: &str,
    _values_only: bool,
    global: &GlobalOptions,
) -> Result<()> {
    if global.dry_run {
        println!("Would clear range {} in {}", range, sheet);
        return Ok(());
    }

    let mut workbook = Workbook::open(file)?;
    let range_ref = Range::parse(range)?;

    for cell_ref in range_ref.cells() {
        workbook.clear_cell(sheet, &cell_ref)?;
    }

    workbook.save()?;

    if !global.quiet {
        println!("Cleared range {}", range.cyan());
    }

    Ok(())
}

fn fill(
    file: &std::path::Path,
    sheet: &str,
    range: &str,
    value: &str,
    global: &GlobalOptions,
) -> Result<()> {
    if global.dry_run {
        println!("Would fill range {} with '{}' in {}", range, value, sheet);
        return Ok(());
    }

    let mut workbook = Workbook::open(file)?;
    let range_ref = Range::parse(range)?;
    let cell_value = super::cell::parse_auto_value(value);

    for cell_ref in range_ref.cells() {
        workbook.set_cell(sheet, cell_ref, cell_value.clone())?;
    }

    workbook.save()?;

    if !global.quiet {
        let count = range_ref.cell_count();
        println!(
            "Filled {} cells in {}",
            count.to_string().green(),
            range.cyan()
        );
    }

    Ok(())
}

fn merge(file: &std::path::Path, sheet: &str, range: &str, global: &GlobalOptions) -> Result<()> {
    if global.dry_run {
        println!("Would merge range {} in {}", range, sheet);
        return Ok(());
    }

    let mut workbook = Workbook::open(file)?;
    let range_ref = Range::parse(range)?;

    let sheet_obj =
        workbook
            .get_sheet_mut(sheet)
            .ok_or_else(|| xlex_core::XlexError::SheetNotFound {
                name: sheet.to_string(),
            })?;

    sheet_obj.add_merged_range(range_ref);
    let _ = sheet_obj;
    workbook.save()?;

    if !global.quiet {
        println!("Merged range {}", range.cyan());
    }

    Ok(())
}

fn unmerge(file: &std::path::Path, sheet: &str, range: &str, global: &GlobalOptions) -> Result<()> {
    if global.dry_run {
        println!("Would unmerge range {} in {}", range, sheet);
        return Ok(());
    }

    let mut workbook = Workbook::open(file)?;
    let range_ref = Range::parse(range)?;

    let sheet_obj =
        workbook
            .get_sheet_mut(sheet)
            .ok_or_else(|| xlex_core::XlexError::SheetNotFound {
                name: sheet.to_string(),
            })?;

    sheet_obj.remove_merged_range(&range_ref);
    let _ = sheet_obj;
    workbook.save()?;

    if !global.quiet {
        println!("Unmerged range {}", range.cyan());
    }

    Ok(())
}

fn name(
    file: &std::path::Path,
    name: &str,
    range: &str,
    sheet: Option<&str>,
    global: &GlobalOptions,
) -> Result<()> {
    if global.dry_run {
        println!("Would define named range '{}' as {}", name, range);
        return Ok(());
    }

    let mut workbook = Workbook::open(file)?;

    // Validate range syntax
    let _range_ref = Range::parse(range)?;

    // Build the reference with sheet name if provided
    let reference = if let Some(sheet_name) = sheet {
        // Verify sheet exists
        if workbook.get_sheet(sheet_name).is_none() {
            anyhow::bail!("Sheet '{}' not found", sheet_name);
        }
        format!("'{}'!{}", sheet_name, range)
    } else {
        range.to_string()
    };

    // Get sheet index for local scope
    let local_sheet_id = sheet.and_then(|s| workbook.sheet_names().iter().position(|n| *n == s));

    let defined_name = DefinedName {
        name: name.to_string(),
        reference,
        local_sheet_id,
        comment: None,
        hidden: false,
    };

    workbook.set_defined_name(defined_name);
    workbook.save()?;

    if !global.quiet {
        if global.format == OutputFormat::Json {
            let json = serde_json::json!({
                "name": name,
                "range": range,
                "sheet": sheet,
            });
            println!("{}", serde_json::to_string_pretty(&json)?);
        } else {
            println!(
                "{} Defined named range '{}' as {}",
                "✓".green(),
                name.cyan(),
                range.yellow()
            );
        }
    }

    Ok(())
}

fn names(file: &std::path::Path, global: &GlobalOptions) -> Result<()> {
    let workbook = Workbook::open(file)?;
    let defined_names = workbook.defined_names();

    if global.format == OutputFormat::Json {
        let names: Vec<serde_json::Value> = defined_names
            .iter()
            .map(|dn| {
                serde_json::json!({
                    "name": dn.name,
                    "reference": dn.reference,
                    "localSheetId": dn.local_sheet_id,
                    "hidden": dn.hidden,
                    "comment": dn.comment,
                })
            })
            .collect();
        println!("{}", serde_json::to_string_pretty(&names)?);
    } else if defined_names.is_empty() {
        println!("No named ranges defined");
    } else {
        println!("{} named range(s):", defined_names.len());
        for dn in defined_names {
            let scope = if let Some(idx) = dn.local_sheet_id {
                format!(" (sheet {})", idx)
            } else {
                " (global)".to_string()
            };
            println!(
                "  {} {} {}{}",
                dn.name.cyan(),
                "→".dimmed(),
                dn.reference.yellow(),
                scope.dimmed()
            );
        }
    }

    Ok(())
}

fn validate(
    file: &std::path::Path,
    sheet: &str,
    range: &str,
    rule: &str,
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

    let mut valid = true;
    let mut errors: Vec<String> = Vec::new();

    for cell_ref in range_ref.cells() {
        let value = sheet_obj.get_value(&cell_ref);

        match rule {
            "nonempty" => {
                if value.is_empty() {
                    valid = false;
                    errors.push(format!("{} is empty", cell_ref.to_a1()));
                }
            }
            "numeric" => {
                if !matches!(value, xlex_core::CellValue::Number(_)) && !value.is_empty() {
                    valid = false;
                    errors.push(format!("{} is not numeric", cell_ref.to_a1()));
                }
            }
            _ => {
                anyhow::bail!("Unknown validation rule: {}", rule);
            }
        }
    }

    if global.format == OutputFormat::Json {
        let json = serde_json::json!({
            "valid": valid,
            "range": range,
            "rule": rule,
            "errors": errors,
        });
        println!("{}", serde_json::to_string_pretty(&json)?);
    } else {
        if valid {
            println!("{}: Validation passed", "✓".green());
        } else {
            println!("{}: Validation failed", "✗".red());
            for err in &errors {
                println!("  - {}", err);
            }
        }
    }

    Ok(())
}

fn sort(
    file: &std::path::Path,
    sheet: &str,
    range: &str,
    column: Option<&str>,
    descending: bool,
    global: &GlobalOptions,
) -> Result<()> {
    if global.dry_run {
        println!("Would sort range {} in {}", range, sheet);
        return Ok(());
    }

    let mut workbook = Workbook::open(file)?;
    let range_ref = Range::parse(range)?;

    // Determine sort column (relative to range start)
    let sort_col = match column {
        Some(c) => {
            let col_num = xlex_core::CellRef::col_from_letters_pub(c)
                .ok_or_else(|| anyhow::anyhow!("Invalid column: {}", c))?;
            if col_num < range_ref.start.col || col_num > range_ref.end.col {
                anyhow::bail!("Sort column {} is outside range {}", c, range);
            }
            col_num - range_ref.start.col
        }
        None => 0, // First column
    };

    let sheet_obj =
        workbook
            .get_sheet(sheet)
            .ok_or_else(|| xlex_core::XlexError::SheetNotFound {
                name: sheet.to_string(),
            })?;

    // Collect all rows as tuples of (sort_key, row_data)
    let mut rows: Vec<(xlex_core::CellValue, Vec<(u32, xlex_core::CellValue)>)> = Vec::new();

    for row in range_ref.start.row..=range_ref.end.row {
        let mut row_data: Vec<(u32, xlex_core::CellValue)> = Vec::new();
        let mut sort_key = xlex_core::CellValue::Empty;

        for col in range_ref.start.col..=range_ref.end.col {
            let cell_ref = xlex_core::CellRef::new(col, row);
            let value = sheet_obj.get_value(&cell_ref);

            if col - range_ref.start.col == sort_col {
                sort_key = value.clone();
            }

            row_data.push((col - range_ref.start.col, value));
        }

        rows.push((sort_key, row_data));
    }

    let _ = sheet_obj;

    // Sort rows
    rows.sort_by(|a, b| {
        let cmp = compare_cell_values(&a.0, &b.0);
        if descending {
            cmp.reverse()
        } else {
            cmp
        }
    });

    // Write back sorted data
    let sheet_obj =
        workbook
            .get_sheet_mut(sheet)
            .ok_or_else(|| xlex_core::XlexError::SheetNotFound {
                name: sheet.to_string(),
            })?;

    for (row_idx, (_, row_data)) in rows.iter().enumerate() {
        let target_row = range_ref.start.row + row_idx as u32;
        for (col_offset, value) in row_data {
            let cell_ref = xlex_core::CellRef::new(range_ref.start.col + col_offset, target_row);
            if value.is_empty() {
                sheet_obj.clear_cell(&cell_ref);
            } else {
                sheet_obj.set_cell(cell_ref, value.clone());
            }
        }
    }

    let _ = sheet_obj;
    workbook.save()?;

    if !global.quiet {
        let order = if descending {
            "descending"
        } else {
            "ascending"
        };
        println!("Sorted range {} ({})", range.cyan(), order);
    }

    Ok(())
}

/// Compare two cell values for sorting
fn compare_cell_values(a: &xlex_core::CellValue, b: &xlex_core::CellValue) -> std::cmp::Ordering {
    use std::cmp::Ordering;
    use xlex_core::CellValue;

    match (a, b) {
        (CellValue::Empty, CellValue::Empty) => Ordering::Equal,
        (CellValue::Empty, _) => Ordering::Greater, // Empty values go last
        (_, CellValue::Empty) => Ordering::Less,
        (CellValue::Number(n1), CellValue::Number(n2)) => {
            n1.partial_cmp(n2).unwrap_or(Ordering::Equal)
        }
        (CellValue::String(s1), CellValue::String(s2)) => s1.cmp(s2),
        (CellValue::Boolean(b1), CellValue::Boolean(b2)) => b1.cmp(b2),
        // Mixed types: numbers < booleans < strings
        (CellValue::Number(_), _) => Ordering::Less,
        (_, CellValue::Number(_)) => Ordering::Greater,
        (CellValue::Boolean(_), CellValue::String(_)) => Ordering::Less,
        (CellValue::String(_), CellValue::Boolean(_)) => Ordering::Greater,
        _ => a.to_display_string().cmp(&b.to_display_string()),
    }
}

fn filter(
    file: &std::path::Path,
    sheet: &str,
    range: &str,
    column: &str,
    value: &str,
    global: &GlobalOptions,
) -> Result<()> {
    let workbook = Workbook::open(file)?;
    let range_ref = Range::parse(range)?;

    // Parse filter column
    let filter_col = xlex_core::CellRef::col_from_letters_pub(column)
        .ok_or_else(|| anyhow::anyhow!("Invalid column: {}", column))?;
    if filter_col < range_ref.start.col || filter_col > range_ref.end.col {
        anyhow::bail!("Filter column {} is outside range {}", column, range);
    }

    let sheet_obj =
        workbook
            .get_sheet(sheet)
            .ok_or_else(|| xlex_core::XlexError::SheetNotFound {
                name: sheet.to_string(),
            })?;

    // Find rows matching the filter
    let mut matching_rows: Vec<u32> = Vec::new();

    for row in range_ref.start.row..=range_ref.end.row {
        let cell_ref = xlex_core::CellRef::new(filter_col, row);
        let cell_value = sheet_obj.get_value(&cell_ref);
        let cell_str = cell_value.to_display_string();

        if cell_str.contains(value) {
            matching_rows.push(row);
        }
    }

    // Output matching rows
    if global.format == OutputFormat::Json {
        let mut result_rows: Vec<Vec<serde_json::Value>> = Vec::new();

        for row in &matching_rows {
            let mut row_values: Vec<serde_json::Value> = Vec::new();
            for col in range_ref.start.col..=range_ref.end.col {
                let cell_ref = xlex_core::CellRef::new(col, *row);
                let val = sheet_obj.get_value(&cell_ref);
                row_values.push(match val {
                    xlex_core::CellValue::Empty => serde_json::Value::Null,
                    xlex_core::CellValue::String(s) => serde_json::Value::String(s),
                    xlex_core::CellValue::Number(n) => serde_json::json!(n),
                    xlex_core::CellValue::Boolean(b) => serde_json::Value::Bool(b),
                    _ => serde_json::Value::String(val.to_display_string()),
                });
            }
            result_rows.push(row_values);
        }

        let json = serde_json::json!({
            "filter": { "column": column, "value": value },
            "matches": matching_rows.len(),
            "rows": result_rows,
        });
        println!("{}", serde_json::to_string_pretty(&json)?);
    } else {
        println!(
            "Found {} matching rows:",
            matching_rows.len().to_string().green()
        );
        for row in &matching_rows {
            let mut values: Vec<String> = Vec::new();
            for col in range_ref.start.col..=range_ref.end.col {
                let cell_ref = xlex_core::CellRef::new(col, *row);
                let val = sheet_obj.get_value(&cell_ref);
                values.push(val.to_display_string());
            }
            println!(
                "  {}: {}",
                format!("Row {}", row).cyan(),
                values.join(" | ")
            );
        }
    }

    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;
    use tempfile::TempDir;
    use xlex_core::CellValue;

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

    fn setup_test_data(file: &std::path::Path) {
        let mut wb = Workbook::open(file).unwrap();
        // Set up test data in A1:C3 (1-indexed: col 1-3, row 1-3)
        for row in 1..=3u32 {
            for col in 1..=3u32 {
                let cell_ref = xlex_core::CellRef::new(col, row);
                let value = CellValue::Number((row * 10 + col - 1) as f64);
                wb.set_cell("Sheet1", cell_ref, value).unwrap();
            }
        }
        wb.save().unwrap();
    }

    #[test]
    fn test_get_range() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "get.xlsx");
        setup_test_data(&file_path);

        let result = get(&file_path, "Sheet1", "A1:C3", &default_global());
        assert!(result.is_ok());
    }

    #[test]
    fn test_get_range_json() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "get_json.xlsx");
        setup_test_data(&file_path);

        let mut global = default_global();
        global.format = OutputFormat::Json;

        let result = get(&file_path, "Sheet1", "A1:C3", &global);
        assert!(result.is_ok());
    }

    #[test]
    fn test_get_range_csv() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "get_csv.xlsx");
        setup_test_data(&file_path);

        let mut global = default_global();
        global.format = OutputFormat::Csv;

        let result = get(&file_path, "Sheet1", "A1:C3", &global);
        assert!(result.is_ok());
    }

    #[test]
    fn test_copy_range() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "copy.xlsx");
        setup_test_data(&file_path);

        let result = copy(&file_path, "Sheet1", "A1:B2", "E1", &default_global());
        assert!(result.is_ok());

        // Verify copy
        let wb = Workbook::open(&file_path).unwrap();
        let cell_ref = xlex_core::CellRef::new(5, 1); // E1 = col 5, row 1 (1-indexed)
        let value = wb.get_cell("Sheet1", &cell_ref).unwrap();
        assert_eq!(value, CellValue::Number(10.0));
    }

    #[test]
    fn test_move_range() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "move.xlsx");
        setup_test_data(&file_path);

        let result = move_range(&file_path, "Sheet1", "A1:B2", "E1", &default_global());
        assert!(result.is_ok());

        // Verify move - source should be empty
        let wb = Workbook::open(&file_path).unwrap();
        let source_cell = xlex_core::CellRef::new(1, 1); // A1 = col 1, row 1 (1-indexed)
        let source_value = wb.get_cell("Sheet1", &source_cell).unwrap();
        assert_eq!(source_value, CellValue::Empty);

        // Destination should have the value
        let dest_cell = xlex_core::CellRef::new(5, 1); // E1 = col 5, row 1 (1-indexed)
        let dest_value = wb.get_cell("Sheet1", &dest_cell).unwrap();
        assert_eq!(dest_value, CellValue::Number(10.0));
    }

    #[test]
    fn test_clear_range() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "clear.xlsx");
        setup_test_data(&file_path);

        let result = clear(&file_path, "Sheet1", "A1:B2", false, &default_global());
        assert!(result.is_ok());

        // Verify cleared
        let wb = Workbook::open(&file_path).unwrap();
        let cell_ref = xlex_core::CellRef::new(1, 1); // A1 = col 1, row 1 (1-indexed)
        let value = wb.get_cell("Sheet1", &cell_ref).unwrap();
        assert_eq!(value, CellValue::Empty);
    }

    #[test]
    fn test_fill_range() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "fill.xlsx");

        let result = fill(&file_path, "Sheet1", "A1:C3", "test", &default_global());
        assert!(result.is_ok());

        // Verify fill
        let wb = Workbook::open(&file_path).unwrap();
        let cell_ref = xlex_core::CellRef::new(1, 1); // A1 = col 1, row 1 (1-indexed)
        let value = wb.get_cell("Sheet1", &cell_ref).unwrap();
        assert_eq!(value, CellValue::String("test".to_string()));
    }

    #[test]
    fn test_merge_range() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "merge.xlsx");

        let result = merge(&file_path, "Sheet1", "A1:B2", &default_global());
        assert!(result.is_ok());
    }

    #[test]
    fn test_unmerge_range() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "unmerge.xlsx");

        // Merge first
        merge(&file_path, "Sheet1", "A1:B2", &default_global()).unwrap();

        let result = unmerge(&file_path, "Sheet1", "A1:B2", &default_global());
        assert!(result.is_ok());
    }

    #[test]
    fn test_name_range() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "name.xlsx");

        let result = name(&file_path, "TestRange", "A1:B10", None, &default_global());
        assert!(result.is_ok());

        let wb = Workbook::open(&file_path).unwrap();
        let names = wb.defined_names();
        assert!(names.iter().any(|n| n.name == "TestRange"));
    }

    #[test]
    fn test_name_range_with_sheet() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "name_sheet.xlsx");

        let result = name(
            &file_path,
            "LocalRange",
            "A1:B10",
            Some("Sheet1"),
            &default_global(),
        );
        assert!(result.is_ok());
    }

    #[test]
    fn test_names_list() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "names.xlsx");

        // Add a named range first
        name(&file_path, "MyRange", "A1:A10", None, &default_global()).unwrap();

        let result = names(&file_path, &default_global());
        assert!(result.is_ok());
    }

    #[test]
    fn test_validate_nonempty() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "validate.xlsx");
        setup_test_data(&file_path);

        let result = validate(&file_path, "Sheet1", "A1:C3", "nonempty", &default_global());
        assert!(result.is_ok());
    }

    #[test]
    fn test_validate_numeric() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "validate_num.xlsx");
        setup_test_data(&file_path);

        let result = validate(&file_path, "Sheet1", "A1:C3", "numeric", &default_global());
        assert!(result.is_ok());
    }

    #[test]
    fn test_sort_range() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "sort.xlsx");
        setup_test_data(&file_path);

        let result = sort(
            &file_path,
            "Sheet1",
            "A1:C3",
            None,
            false,
            &default_global(),
        );
        assert!(result.is_ok());
    }

    #[test]
    fn test_sort_range_descending() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "sort_desc.xlsx");
        setup_test_data(&file_path);

        let result = sort(&file_path, "Sheet1", "A1:C3", None, true, &default_global());
        assert!(result.is_ok());
    }

    #[test]
    fn test_filter_range() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "filter.xlsx");
        setup_test_data(&file_path);

        let result = filter(&file_path, "Sheet1", "A1:C3", "A", "1", &default_global());
        assert!(result.is_ok());
    }

    #[test]
    fn test_range_style() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "style.xlsx");

        let opts = RangeStyleOpts {
            bold: true,
            italic: false,
            ..Default::default()
        };

        let result = range_style(&file_path, "Sheet1", "A1:B2", opts, &default_global());
        assert!(result.is_ok());
    }

    #[test]
    fn test_range_border() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "border.xlsx");

        let opts = RangeBorderOpts {
            all: true,
            style: "thin".to_string(),
            ..Default::default()
        };

        let result = range_border(&file_path, "Sheet1", "A1:B2", opts, &default_global());
        assert!(result.is_ok());
    }

    #[test]
    fn test_compare_cell_values() {
        // Numbers
        assert_eq!(
            compare_cell_values(&CellValue::Number(1.0), &CellValue::Number(2.0)),
            std::cmp::Ordering::Less
        );
        assert_eq!(
            compare_cell_values(&CellValue::Number(2.0), &CellValue::Number(1.0)),
            std::cmp::Ordering::Greater
        );

        // Empty values go last
        assert_eq!(
            compare_cell_values(&CellValue::Empty, &CellValue::Number(1.0)),
            std::cmp::Ordering::Greater
        );

        // Strings
        assert_eq!(
            compare_cell_values(
                &CellValue::String("a".to_string()),
                &CellValue::String("b".to_string())
            ),
            std::cmp::Ordering::Less
        );
    }

    #[test]
    fn test_dry_run_operations() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "dry.xlsx");

        let mut global = default_global();
        global.dry_run = true;

        // All these should succeed without modifying the file
        assert!(copy(&file_path, "Sheet1", "A1:B2", "E1", &global).is_ok());
        assert!(move_range(&file_path, "Sheet1", "A1:B2", "E1", &global).is_ok());
        assert!(clear(&file_path, "Sheet1", "A1:B2", false, &global).is_ok());
        assert!(fill(&file_path, "Sheet1", "A1:B2", "test", &global).is_ok());
        assert!(merge(&file_path, "Sheet1", "A1:B2", &global).is_ok());
        assert!(unmerge(&file_path, "Sheet1", "A1:B2", &global).is_ok());
    }

    // Additional tests for better coverage

    #[test]
    fn test_run_get_command() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "run_get.xlsx");
        setup_test_data(&file_path);

        let args = RangeArgs {
            command: RangeCommand::Get {
                file: file_path,
                sheet: "Sheet1".to_string(),
                range: "A1:C3".to_string(),
            },
        };

        let result = run(&args, &default_global());
        assert!(result.is_ok());
    }

    #[test]
    fn test_run_copy_command() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "run_copy.xlsx");
        setup_test_data(&file_path);

        let args = RangeArgs {
            command: RangeCommand::Copy {
                file: file_path,
                sheet: "Sheet1".to_string(),
                source: "A1:B2".to_string(),
                dest: "E1".to_string(),
            },
        };

        let result = run(&args, &default_global());
        assert!(result.is_ok());
    }

    #[test]
    fn test_run_move_command() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "run_move.xlsx");
        setup_test_data(&file_path);

        let args = RangeArgs {
            command: RangeCommand::Move {
                file: file_path,
                sheet: "Sheet1".to_string(),
                source: "A1:B2".to_string(),
                dest: "E1".to_string(),
            },
        };

        let result = run(&args, &default_global());
        assert!(result.is_ok());
    }

    #[test]
    fn test_run_clear_command() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "run_clear.xlsx");
        setup_test_data(&file_path);

        let args = RangeArgs {
            command: RangeCommand::Clear {
                file: file_path,
                sheet: "Sheet1".to_string(),
                range: "A1:B2".to_string(),
                values_only: false,
            },
        };

        let result = run(&args, &default_global());
        assert!(result.is_ok());
    }

    #[test]
    fn test_run_fill_command() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "run_fill.xlsx");

        let args = RangeArgs {
            command: RangeCommand::Fill {
                file: file_path,
                sheet: "Sheet1".to_string(),
                range: "A1:B2".to_string(),
                value: "test".to_string(),
            },
        };

        let result = run(&args, &default_global());
        assert!(result.is_ok());
    }

    #[test]
    fn test_run_merge_command() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "run_merge.xlsx");

        let args = RangeArgs {
            command: RangeCommand::Merge {
                file: file_path,
                sheet: "Sheet1".to_string(),
                range: "A1:B2".to_string(),
            },
        };

        let result = run(&args, &default_global());
        assert!(result.is_ok());
    }

    #[test]
    fn test_run_unmerge_command() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "run_unmerge.xlsx");

        merge(&file_path, "Sheet1", "A1:B2", &default_global()).unwrap();

        let args = RangeArgs {
            command: RangeCommand::Unmerge {
                file: file_path,
                sheet: "Sheet1".to_string(),
                range: "A1:B2".to_string(),
            },
        };

        let result = run(&args, &default_global());
        assert!(result.is_ok());
    }

    #[test]
    fn test_copy_json_output() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "copy_json.xlsx");
        setup_test_data(&file_path);

        let mut global = default_global();
        global.format = OutputFormat::Json;
        global.quiet = false;

        let result = copy(&file_path, "Sheet1", "A1:B2", "E1", &global);
        assert!(result.is_ok());
    }

    #[test]
    fn test_move_json_output() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "move_json.xlsx");
        setup_test_data(&file_path);

        let mut global = default_global();
        global.format = OutputFormat::Json;
        global.quiet = false;

        let result = move_range(&file_path, "Sheet1", "A1:B2", "E1", &global);
        assert!(result.is_ok());
    }

    #[test]
    fn test_clear_json_output() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "clear_json.xlsx");
        setup_test_data(&file_path);

        let mut global = default_global();
        global.format = OutputFormat::Json;
        global.quiet = false;

        let result = clear(&file_path, "Sheet1", "A1:B2", false, &global);
        assert!(result.is_ok());
    }

    #[test]
    fn test_fill_json_output() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "fill_json.xlsx");

        let mut global = default_global();
        global.format = OutputFormat::Json;
        global.quiet = false;

        let result = fill(&file_path, "Sheet1", "A1:B2", "test", &global);
        assert!(result.is_ok());
    }

    #[test]
    fn test_merge_json_output() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "merge_json.xlsx");

        let mut global = default_global();
        global.format = OutputFormat::Json;
        global.quiet = false;

        let result = merge(&file_path, "Sheet1", "A1:B2", &global);
        assert!(result.is_ok());
    }

    #[test]
    fn test_names_json_output() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "names_json.xlsx");

        name(&file_path, "MyRange", "A1:A10", None, &default_global()).unwrap();

        let mut global = default_global();
        global.format = OutputFormat::Json;

        let result = names(&file_path, &global);
        assert!(result.is_ok());
    }

    #[test]
    fn test_validate_json_output() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "validate_json.xlsx");
        setup_test_data(&file_path);

        let mut global = default_global();
        global.format = OutputFormat::Json;

        let result = validate(&file_path, "Sheet1", "A1:C3", "nonempty", &global);
        assert!(result.is_ok());
    }

    #[test]
    fn test_sort_json_output() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "sort_json.xlsx");
        setup_test_data(&file_path);

        let mut global = default_global();
        global.format = OutputFormat::Json;
        global.quiet = false;

        let result = sort(&file_path, "Sheet1", "A1:C3", None, false, &global);
        assert!(result.is_ok());
    }

    #[test]
    fn test_filter_json_output() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "filter_json.xlsx");
        setup_test_data(&file_path);

        let mut global = default_global();
        global.format = OutputFormat::Json;

        let result = filter(&file_path, "Sheet1", "A1:C3", "A", "1", &global);
        assert!(result.is_ok());
    }

    #[test]
    fn test_clear_with_styles() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "clear_styles.xlsx");
        setup_test_data(&file_path);

        let result = clear(&file_path, "Sheet1", "A1:B2", true, &default_global());
        assert!(result.is_ok());
    }

    #[test]
    fn test_fill_number_value() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "fill_num.xlsx");

        let result = fill(&file_path, "Sheet1", "A1:C3", "123", &default_global());
        assert!(result.is_ok());

        let wb = Workbook::open(&file_path).unwrap();
        let cell_ref = xlex_core::CellRef::new(1, 1);
        let value = wb.get_cell("Sheet1", &cell_ref).unwrap();
        assert_eq!(value, CellValue::Number(123.0));
    }

    #[test]
    fn test_sort_with_column() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "sort_col.xlsx");
        setup_test_data(&file_path);

        let result = sort(
            &file_path,
            "Sheet1",
            "A1:C3",
            Some("B"),
            false,
            &default_global(),
        );
        assert!(result.is_ok());
    }

    // More coverage tests for run commands

    #[test]
    fn test_run_style_command() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "run_style.xlsx");
        setup_test_data(&file_path);

        let args = RangeArgs {
            command: RangeCommand::Style {
                file: file_path,
                sheet: "Sheet1".to_string(),
                range: "A1:B2".to_string(),
                bold: true,
                italic: true,
                underline: false,
                font: Some("Arial".to_string()),
                font_size: Some(12.0),
                text_color: Some("FF0000".to_string()),
                bg_color: Some("FFFF00".to_string()),
                align: Some("center".to_string()),
                valign: Some("middle".to_string()),
                wrap: true,
                number_format: None,
                percent: false,
                currency: None,
                date_format: None,
            },
        };

        let result = run(&args, &default_global());
        assert!(result.is_ok());
    }

    #[test]
    fn test_run_border_command() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "run_border.xlsx");

        let args = RangeArgs {
            command: RangeCommand::Border {
                file: file_path,
                sheet: "Sheet1".to_string(),
                range: "A1:B2".to_string(),
                all: true,
                outline: false,
                top: false,
                bottom: false,
                left: false,
                right: false,
                none: false,
                style: "thin".to_string(),
                border_color: Some("000000".to_string()),
            },
        };

        let result = run(&args, &default_global());
        assert!(result.is_ok());
    }

    #[test]
    fn test_run_name_command() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "run_name.xlsx");

        let args = RangeArgs {
            command: RangeCommand::Name {
                file: file_path,
                name: "TestRange".to_string(),
                range: "A1:B10".to_string(),
                sheet: None,
            },
        };

        let result = run(&args, &default_global());
        assert!(result.is_ok());
    }

    #[test]
    fn test_run_names_command() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "run_names.xlsx");

        // Add a named range first
        name(&file_path, "MyRange", "A1:A10", None, &default_global()).unwrap();

        let args = RangeArgs {
            command: RangeCommand::Names { file: file_path },
        };

        let result = run(&args, &default_global());
        assert!(result.is_ok());
    }

    #[test]
    fn test_run_validate_command() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "run_validate.xlsx");
        setup_test_data(&file_path);

        let args = RangeArgs {
            command: RangeCommand::Validate {
                file: file_path,
                sheet: "Sheet1".to_string(),
                range: "A1:C3".to_string(),
                rule: "nonempty".to_string(),
            },
        };

        let result = run(&args, &default_global());
        assert!(result.is_ok());
    }

    #[test]
    fn test_run_sort_command() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "run_sort.xlsx");
        setup_test_data(&file_path);

        let args = RangeArgs {
            command: RangeCommand::Sort {
                file: file_path,
                sheet: "Sheet1".to_string(),
                range: "A1:C3".to_string(),
                column: None,
                descending: false,
            },
        };

        let result = run(&args, &default_global());
        assert!(result.is_ok());
    }

    #[test]
    fn test_run_filter_command() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "run_filter.xlsx");
        setup_test_data(&file_path);

        let args = RangeArgs {
            command: RangeCommand::Filter {
                file: file_path,
                sheet: "Sheet1".to_string(),
                range: "A1:C3".to_string(),
                column: "A".to_string(),
                value: "1".to_string(),
            },
        };

        let result = run(&args, &default_global());
        assert!(result.is_ok());
    }

    // Verbose output tests

    #[test]
    fn test_copy_verbose() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "copy_verbose.xlsx");
        setup_test_data(&file_path);

        let mut global = default_global();
        global.verbose = true;
        global.quiet = false;

        let result = copy(&file_path, "Sheet1", "A1:B2", "E1", &global);
        assert!(result.is_ok());
    }

    #[test]
    fn test_move_verbose() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "move_verbose.xlsx");
        setup_test_data(&file_path);

        let mut global = default_global();
        global.verbose = true;
        global.quiet = false;

        let result = move_range(&file_path, "Sheet1", "A1:B2", "E1", &global);
        assert!(result.is_ok());
    }

    #[test]
    fn test_clear_verbose() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "clear_verbose.xlsx");
        setup_test_data(&file_path);

        let mut global = default_global();
        global.verbose = true;
        global.quiet = false;

        let result = clear(&file_path, "Sheet1", "A1:B2", false, &global);
        assert!(result.is_ok());
    }

    #[test]
    fn test_fill_verbose() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "fill_verbose.xlsx");

        let mut global = default_global();
        global.verbose = true;
        global.quiet = false;

        let result = fill(&file_path, "Sheet1", "A1:B2", "test", &global);
        assert!(result.is_ok());
    }

    #[test]
    fn test_merge_verbose() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "merge_verbose.xlsx");

        let mut global = default_global();
        global.verbose = true;
        global.quiet = false;

        let result = merge(&file_path, "Sheet1", "A1:B2", &global);
        assert!(result.is_ok());
    }

    #[test]
    fn test_unmerge_verbose() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "unmerge_verbose.xlsx");

        merge(&file_path, "Sheet1", "A1:B2", &default_global()).unwrap();

        let mut global = default_global();
        global.verbose = true;
        global.quiet = false;

        let result = unmerge(&file_path, "Sheet1", "A1:B2", &global);
        assert!(result.is_ok());
    }

    #[test]
    fn test_name_verbose() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "name_verbose.xlsx");

        let mut global = default_global();
        global.verbose = true;
        global.quiet = false;

        let result = name(&file_path, "TestRange", "A1:B10", None, &global);
        assert!(result.is_ok());
    }

    #[test]
    fn test_sort_verbose() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "sort_verbose.xlsx");
        setup_test_data(&file_path);

        let mut global = default_global();
        global.verbose = true;
        global.quiet = false;

        let result = sort(&file_path, "Sheet1", "A1:C3", None, false, &global);
        assert!(result.is_ok());
    }

    // Dry run tests for more operations

    #[test]
    fn test_name_dry_run() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "name_dry.xlsx");

        let mut global = default_global();
        global.dry_run = true;

        let result = name(&file_path, "TestRange", "A1:B10", None, &global);
        assert!(result.is_ok());
    }

    #[test]
    fn test_sort_dry_run() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "sort_dry.xlsx");
        setup_test_data(&file_path);

        let mut global = default_global();
        global.dry_run = true;

        let result = sort(&file_path, "Sheet1", "A1:C3", None, false, &global);
        assert!(result.is_ok());
    }

    #[test]
    fn test_range_style_dry_run() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "style_dry.xlsx");

        let mut global = default_global();
        global.dry_run = true;

        let opts = RangeStyleOpts {
            bold: true,
            ..Default::default()
        };

        let result = range_style(&file_path, "Sheet1", "A1:B2", opts, &global);
        assert!(result.is_ok());
    }

    #[test]
    fn test_range_border_dry_run() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "border_dry.xlsx");

        let mut global = default_global();
        global.dry_run = true;

        let opts = RangeBorderOpts {
            all: true,
            style: "thin".to_string(),
            ..Default::default()
        };

        let result = range_border(&file_path, "Sheet1", "A1:B2", opts, &global);
        assert!(result.is_ok());
    }

    // Edge cases and error handling

    #[test]
    fn test_validate_unknown_rule() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "validate_err.xlsx");
        setup_test_data(&file_path);

        let result = validate(
            &file_path,
            "Sheet1",
            "A1:C3",
            "unknown_rule",
            &default_global(),
        );
        assert!(result.is_err());
    }

    #[test]
    fn test_filter_invalid_column() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "filter_col_err.xlsx");
        setup_test_data(&file_path);

        // Column Z is outside A1:C3 range
        let result = filter(&file_path, "Sheet1", "A1:C3", "Z", "1", &default_global());
        assert!(result.is_err());
    }

    #[test]
    fn test_sort_invalid_column() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "sort_col_err.xlsx");
        setup_test_data(&file_path);

        // Column Z is outside A1:C3 range
        let result = sort(
            &file_path,
            "Sheet1",
            "A1:C3",
            Some("Z"),
            false,
            &default_global(),
        );
        assert!(result.is_err());
    }

    #[test]
    fn test_get_sheet_not_found() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "get_no_sheet.xlsx");

        let result = get(&file_path, "NonExistentSheet", "A1:C3", &default_global());
        assert!(result.is_err());
    }

    #[test]
    fn test_copy_sheet_not_found() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "copy_no_sheet.xlsx");

        let result = copy(
            &file_path,
            "NonExistentSheet",
            "A1:B2",
            "E1",
            &default_global(),
        );
        assert!(result.is_err());
    }

    #[test]
    fn test_move_sheet_not_found() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "move_no_sheet.xlsx");

        let result = move_range(
            &file_path,
            "NonExistentSheet",
            "A1:B2",
            "E1",
            &default_global(),
        );
        assert!(result.is_err());
    }

    #[test]
    fn test_clear_sheet_not_found() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "clear_no_sheet.xlsx");

        let result = clear(
            &file_path,
            "NonExistentSheet",
            "A1:B2",
            false,
            &default_global(),
        );
        assert!(result.is_err());
    }

    #[test]
    fn test_merge_sheet_not_found() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "merge_no_sheet.xlsx");

        let result = merge(&file_path, "NonExistentSheet", "A1:B2", &default_global());
        assert!(result.is_err());
    }

    #[test]
    fn test_validate_sheet_not_found() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "validate_no_sheet.xlsx");

        let result = validate(
            &file_path,
            "NonExistentSheet",
            "A1:C3",
            "nonempty",
            &default_global(),
        );
        assert!(result.is_err());
    }

    #[test]
    fn test_sort_sheet_not_found() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "sort_no_sheet.xlsx");

        let result = sort(
            &file_path,
            "NonExistentSheet",
            "A1:C3",
            None,
            false,
            &default_global(),
        );
        assert!(result.is_err());
    }

    #[test]
    fn test_filter_sheet_not_found() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "filter_no_sheet.xlsx");

        let result = filter(
            &file_path,
            "NonExistentSheet",
            "A1:C3",
            "A",
            "1",
            &default_global(),
        );
        assert!(result.is_err());
    }

    #[test]
    fn test_name_sheet_not_found() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "name_no_sheet.xlsx");

        let result = name(
            &file_path,
            "TestRange",
            "A1:B10",
            Some("NonExistentSheet"),
            &default_global(),
        );
        assert!(result.is_err());
    }

    #[test]
    fn test_range_style_sheet_not_found() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "style_no_sheet.xlsx");

        let opts = RangeStyleOpts {
            bold: true,
            ..Default::default()
        };

        let result = range_style(
            &file_path,
            "NonExistentSheet",
            "A1:B2",
            opts,
            &default_global(),
        );
        assert!(result.is_err());
    }

    #[test]
    fn test_range_border_sheet_not_found() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "border_no_sheet.xlsx");

        let opts = RangeBorderOpts {
            all: true,
            style: "thin".to_string(),
            ..Default::default()
        };

        let result = range_border(
            &file_path,
            "NonExistentSheet",
            "A1:B2",
            opts,
            &default_global(),
        );
        assert!(result.is_err());
    }

    // Compare cell values edge cases

    #[test]
    fn test_compare_cell_values_booleans() {
        assert_eq!(
            compare_cell_values(&CellValue::Boolean(false), &CellValue::Boolean(true)),
            std::cmp::Ordering::Less
        );
        assert_eq!(
            compare_cell_values(&CellValue::Boolean(true), &CellValue::Boolean(false)),
            std::cmp::Ordering::Greater
        );
        assert_eq!(
            compare_cell_values(&CellValue::Boolean(true), &CellValue::Boolean(true)),
            std::cmp::Ordering::Equal
        );
    }

    #[test]
    fn test_compare_cell_values_mixed_types() {
        // Numbers come before booleans
        assert_eq!(
            compare_cell_values(&CellValue::Number(1.0), &CellValue::Boolean(true)),
            std::cmp::Ordering::Less
        );
        assert_eq!(
            compare_cell_values(&CellValue::Boolean(true), &CellValue::Number(1.0)),
            std::cmp::Ordering::Greater
        );

        // Booleans come before strings
        assert_eq!(
            compare_cell_values(
                &CellValue::Boolean(true),
                &CellValue::String("test".to_string())
            ),
            std::cmp::Ordering::Less
        );
        assert_eq!(
            compare_cell_values(
                &CellValue::String("test".to_string()),
                &CellValue::Boolean(true)
            ),
            std::cmp::Ordering::Greater
        );
    }

    #[test]
    fn test_compare_cell_values_equal_empty() {
        assert_eq!(
            compare_cell_values(&CellValue::Empty, &CellValue::Empty),
            std::cmp::Ordering::Equal
        );
    }

    // Range style various alignments

    #[test]
    fn test_range_style_alignments() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "style_align.xlsx");

        // Test left alignment
        let opts = RangeStyleOpts {
            align: Some("left".to_string()),
            ..Default::default()
        };
        let result = range_style(&file_path, "Sheet1", "A1", opts, &default_global());
        assert!(result.is_ok());

        // Test right alignment
        let opts = RangeStyleOpts {
            align: Some("right".to_string()),
            ..Default::default()
        };
        let result = range_style(&file_path, "Sheet1", "B1", opts, &default_global());
        assert!(result.is_ok());

        // Test justify alignment
        let opts = RangeStyleOpts {
            align: Some("justify".to_string()),
            ..Default::default()
        };
        let result = range_style(&file_path, "Sheet1", "C1", opts, &default_global());
        assert!(result.is_ok());

        // Test unknown alignment (defaults to General)
        let opts = RangeStyleOpts {
            align: Some("unknown".to_string()),
            ..Default::default()
        };
        let result = range_style(&file_path, "Sheet1", "D1", opts, &default_global());
        assert!(result.is_ok());
    }

    #[test]
    fn test_range_style_vertical_alignments() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "style_valign.xlsx");

        // Test top alignment
        let opts = RangeStyleOpts {
            valign: Some("top".to_string()),
            ..Default::default()
        };
        let result = range_style(&file_path, "Sheet1", "A1", opts, &default_global());
        assert!(result.is_ok());

        // Test center alignment
        let opts = RangeStyleOpts {
            valign: Some("center".to_string()),
            ..Default::default()
        };
        let result = range_style(&file_path, "Sheet1", "B1", opts, &default_global());
        assert!(result.is_ok());

        // Test bottom alignment
        let opts = RangeStyleOpts {
            valign: Some("bottom".to_string()),
            ..Default::default()
        };
        let result = range_style(&file_path, "Sheet1", "C1", opts, &default_global());
        assert!(result.is_ok());

        // Test unknown alignment (defaults to Center)
        let opts = RangeStyleOpts {
            valign: Some("unknown".to_string()),
            ..Default::default()
        };
        let result = range_style(&file_path, "Sheet1", "D1", opts, &default_global());
        assert!(result.is_ok());
    }

    #[test]
    fn test_range_style_number_formats() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "style_numfmt.xlsx");

        // Test percent format
        let opts = RangeStyleOpts {
            percent: true,
            ..Default::default()
        };
        let result = range_style(&file_path, "Sheet1", "A1", opts, &default_global());
        assert!(result.is_ok());

        // Test currency format
        let opts = RangeStyleOpts {
            currency: Some("USD".to_string()),
            ..Default::default()
        };
        let result = range_style(&file_path, "Sheet1", "B1", opts, &default_global());
        assert!(result.is_ok());

        // Test date format
        let opts = RangeStyleOpts {
            date_format: Some("YYYY-MM-DD".to_string()),
            ..Default::default()
        };
        let result = range_style(&file_path, "Sheet1", "C1", opts, &default_global());
        assert!(result.is_ok());

        // Test custom number format
        let opts = RangeStyleOpts {
            number_format: Some("#,##0.00".to_string()),
            ..Default::default()
        };
        let result = range_style(&file_path, "Sheet1", "D1", opts, &default_global());
        assert!(result.is_ok());
    }

    // Range border various styles

    #[test]
    fn test_range_border_styles() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "border_styles.xlsx");

        let styles = [
            "thin", "medium", "thick", "dashed", "dotted", "double", "hair", "unknown",
        ];
        for (i, style) in styles.iter().enumerate() {
            let opts = RangeBorderOpts {
                all: true,
                style: style.to_string(),
                ..Default::default()
            };
            let range = format!("A{}:B{}", i + 1, i + 1);
            let result = range_border(&file_path, "Sheet1", &range, opts, &default_global());
            assert!(result.is_ok());
        }
    }

    #[test]
    fn test_range_border_outline() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "border_outline.xlsx");

        let opts = RangeBorderOpts {
            outline: true,
            style: "medium".to_string(),
            ..Default::default()
        };

        let result = range_border(&file_path, "Sheet1", "A1:C3", opts, &default_global());
        assert!(result.is_ok());
    }

    #[test]
    fn test_range_border_individual() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "border_individual.xlsx");

        let opts = RangeBorderOpts {
            top: true,
            bottom: true,
            left: true,
            right: true,
            style: "thin".to_string(),
            ..Default::default()
        };

        let result = range_border(&file_path, "Sheet1", "A1:C3", opts, &default_global());
        assert!(result.is_ok());
    }

    #[test]
    fn test_range_border_remove() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "border_remove.xlsx");

        // First add borders
        let add_opts = RangeBorderOpts {
            all: true,
            style: "thin".to_string(),
            ..Default::default()
        };
        range_border(&file_path, "Sheet1", "A1:C3", add_opts, &default_global()).unwrap();

        // Then remove borders
        let remove_opts = RangeBorderOpts {
            none: true,
            style: "thin".to_string(),
            ..Default::default()
        };

        let mut global = default_global();
        global.quiet = false;

        let result = range_border(&file_path, "Sheet1", "A1:C3", remove_opts, &global);
        assert!(result.is_ok());
    }

    // Names list with various states

    #[test]
    fn test_names_empty() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "names_empty.xlsx");

        let result = names(&file_path, &default_global());
        assert!(result.is_ok());
    }

    #[test]
    fn test_name_json_output() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "name_json_out.xlsx");

        let mut global = default_global();
        global.format = OutputFormat::Json;
        global.quiet = false;

        let result = name(&file_path, "TestRange", "A1:B10", None, &global);
        assert!(result.is_ok());
    }

    // Validation edge cases

    #[test]
    fn test_validate_nonempty_fails() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "validate_fail.xlsx");

        // Don't set up test data, so cells are empty
        let result = validate(&file_path, "Sheet1", "A1:C3", "nonempty", &default_global());
        assert!(result.is_ok()); // Should succeed but report validation failures
    }

    #[test]
    fn test_validate_numeric_fails() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "validate_num_fail.xlsx");

        // Set string values
        {
            let mut wb = Workbook::open(&file_path).unwrap();
            wb.set_cell(
                "Sheet1",
                xlex_core::CellRef::new(1, 1),
                CellValue::String("text".to_string()),
            )
            .unwrap();
            wb.save().unwrap();
        }

        let result = validate(&file_path, "Sheet1", "A1:A1", "numeric", &default_global());
        assert!(result.is_ok()); // Should succeed but report validation failures
    }

    // Get with text output verbose

    #[test]
    fn test_get_range_text_output() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "get_text.xlsx");
        setup_test_data(&file_path);

        let mut global = default_global();
        global.quiet = false;

        let result = get(&file_path, "Sheet1", "A1:C3", &global);
        assert!(result.is_ok());
    }

    // Filter verbose output

    #[test]
    fn test_filter_text_output() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = create_test_workbook(&temp_dir, "filter_text.xlsx");
        setup_test_data(&file_path);

        let mut global = default_global();
        global.quiet = false;

        let result = filter(&file_path, "Sheet1", "A1:C3", "A", "1", &global);
        assert!(result.is_ok());
    }
}
