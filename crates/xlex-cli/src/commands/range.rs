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
        color: Option<String>,
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
            color,
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
                color: color.clone(),
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
        let sheet_obj = workbook
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
            
            let sheet_obj = workbook
                .get_sheet_mut(sheet)
                .ok_or_else(|| xlex_core::XlexError::SheetNotFound {
                    name: sheet.to_string(),
                })?;
            sheet_obj.set_cell_style(&cell_ref, Some(style_id));
        }
    }

    workbook.save()?;

    if !global.quiet {
        let action = if opts.none { "Removed" } else { "Applied" };
        println!("{} {} borders on range {}", "✓".green(), action, range.cyan());
    }
    Ok(())
}

fn get(
    file: &std::path::Path,
    sheet: &str,
    range: &str,
    global: &GlobalOptions,
) -> Result<()> {
    let workbook = Workbook::open(file)?;
    let range_ref = Range::parse(range)?;
    let sheet_obj = workbook
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
            println!("{}: {}", format!("Row {}", row_num).cyan(), values.join(" | "));
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

    let sheet_obj = workbook
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
    let sheet_obj = workbook
        .get_sheet_mut(sheet)
        .ok_or_else(|| xlex_core::XlexError::SheetNotFound {
            name: sheet.to_string(),
        })?;

    for (col_offset, row_offset, value) in values {
        let dest_ref = xlex_core::CellRef::new(
            dest_cell.col + col_offset,
            dest_cell.row + row_offset,
        );
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

    let sheet_obj = workbook
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

    let sheet_obj = workbook
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
            let dest_ref = xlex_core::CellRef::new(
                dest_cell.col + col_offset,
                dest_cell.row + row_offset,
            );
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
        println!("Filled {} cells in {}", count.to_string().green(), range.cyan());
    }

    Ok(())
}

fn merge(
    file: &std::path::Path,
    sheet: &str,
    range: &str,
    global: &GlobalOptions,
) -> Result<()> {
    if global.dry_run {
        println!("Would merge range {} in {}", range, sheet);
        return Ok(());
    }

    let mut workbook = Workbook::open(file)?;
    let range_ref = Range::parse(range)?;

    let sheet_obj = workbook
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

fn unmerge(
    file: &std::path::Path,
    sheet: &str,
    range: &str,
    global: &GlobalOptions,
) -> Result<()> {
    if global.dry_run {
        println!("Would unmerge range {} in {}", range, sheet);
        return Ok(());
    }

    let mut workbook = Workbook::open(file)?;
    let range_ref = Range::parse(range)?;

    let sheet_obj = workbook
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
    let local_sheet_id = sheet.and_then(|s| {
        workbook
            .sheet_names()
            .iter()
            .position(|n| *n == s)
    });

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
    let sheet_obj = workbook
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
                if !matches!(value, xlex_core::CellValue::Number(_))
                    && !value.is_empty()
                {
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

    let sheet_obj = workbook
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
    let sheet_obj = workbook
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
        let order = if descending { "descending" } else { "ascending" };
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

    let sheet_obj = workbook
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
        println!("Found {} matching rows:", matching_rows.len().to_string().green());
        for row in &matching_rows {
            let mut values: Vec<String> = Vec::new();
            for col in range_ref.start.col..=range_ref.end.col {
                let cell_ref = xlex_core::CellRef::new(col, *row);
                let val = sheet_obj.get_value(&cell_ref);
                values.push(val.to_display_string());
            }
            println!("  {}: {}", format!("Row {}", row).cyan(), values.join(" | "));
        }
    }

    Ok(())
}
