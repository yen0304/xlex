//! Style types and operations.

use serde::{Deserialize, Serialize};
use std::collections::HashMap;

/// Text alignment options.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default, Serialize, Deserialize)]
pub enum HorizontalAlignment {
    #[default]
    General,
    Left,
    Center,
    Right,
    Fill,
    Justify,
    CenterContinuous,
    Distributed,
}

/// Vertical alignment options.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default, Serialize, Deserialize)]
pub enum VerticalAlignment {
    Top,
    #[default]
    Center,
    Bottom,
    Justify,
    Distributed,
}

/// Border style.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default, Serialize, Deserialize)]
pub enum BorderStyle {
    #[default]
    None,
    Thin,
    Medium,
    Thick,
    Dashed,
    Dotted,
    Double,
    Hair,
    MediumDashed,
    DashDot,
    MediumDashDot,
    DashDotDot,
    MediumDashDotDot,
    SlantDashDot,
}

/// A color value.
#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize, Default)]
pub enum Color {
    /// RGB color (0xRRGGBB)
    Rgb(u32),
    /// Theme color index
    Theme(u32),
    /// Indexed color
    Indexed(u32),
    /// Auto color (black for text, white for background)
    #[default]
    Auto,
}

impl Color {
    /// Creates a new RGB color.
    pub fn rgb(r: u8, g: u8, b: u8) -> Self {
        Self::Rgb(((r as u32) << 16) | ((g as u32) << 8) | (b as u32))
    }

    /// Creates a color from a hex string (e.g., "FF0000" or "#FF0000").
    pub fn from_hex(s: &str) -> Option<Self> {
        let s = s.trim_start_matches('#');
        if s.len() == 6 {
            u32::from_str_radix(s, 16).ok().map(Self::Rgb)
        } else if s.len() == 8 {
            // ARGB format, ignore alpha
            u32::from_str_radix(&s[2..], 16).ok().map(Self::Rgb)
        } else {
            None
        }
    }

    /// Returns the RGB value if this is an RGB color.
    pub fn to_rgb(&self) -> Option<(u8, u8, u8)> {
        match self {
            Self::Rgb(val) => Some((
                ((val >> 16) & 0xFF) as u8,
                ((val >> 8) & 0xFF) as u8,
                (val & 0xFF) as u8,
            )),
            _ => None,
        }
    }

    /// Returns the hex string representation.
    pub fn to_hex(&self) -> Option<String> {
        match self {
            Self::Rgb(val) => Some(format!("{:06X}", val)),
            _ => None,
        }
    }
}

/// Font style.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize, Default)]
pub struct Font {
    /// Font name (e.g., "Arial", "Calibri")
    pub name: Option<String>,
    /// Font size in points
    pub size: Option<f64>,
    /// Bold
    pub bold: bool,
    /// Italic
    pub italic: bool,
    /// Underline
    pub underline: bool,
    /// Strikethrough
    pub strikethrough: bool,
    /// Font color
    pub color: Option<Color>,
}

/// Fill pattern type.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default, Serialize, Deserialize)]
pub enum FillPattern {
    #[default]
    None,
    Solid,
    MediumGray,
    DarkGray,
    LightGray,
    DarkHorizontal,
    DarkVertical,
    DarkDown,
    DarkUp,
    DarkGrid,
    DarkTrellis,
    LightHorizontal,
    LightVertical,
    LightDown,
    LightUp,
    LightGrid,
    LightTrellis,
    Gray125,
    Gray0625,
}

/// Fill style.
#[derive(Debug, Clone, PartialEq, Default, Serialize, Deserialize)]
pub struct Fill {
    /// Fill pattern
    pub pattern: FillPattern,
    /// Foreground color (pattern color)
    pub fg_color: Option<Color>,
    /// Background color
    pub bg_color: Option<Color>,
}

/// Border definition for one side.
#[derive(Debug, Clone, PartialEq, Default, Serialize, Deserialize)]
pub struct BorderSide {
    /// Border style
    pub style: BorderStyle,
    /// Border color
    pub color: Option<Color>,
}

/// Complete border definition.
#[derive(Debug, Clone, PartialEq, Default, Serialize, Deserialize)]
pub struct Border {
    pub left: BorderSide,
    pub right: BorderSide,
    pub top: BorderSide,
    pub bottom: BorderSide,
    pub diagonal: BorderSide,
    pub diagonal_up: bool,
    pub diagonal_down: bool,
}

impl Border {
    /// Creates a border with all sides set to the same style.
    pub fn all(style: BorderStyle, color: Option<Color>) -> Self {
        let side = BorderSide {
            style,
            color: color.clone(),
        };
        Self {
            left: side.clone(),
            right: side.clone(),
            top: side.clone(),
            bottom: side,
            ..Default::default()
        }
    }
}

/// Number format.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct NumberFormat {
    /// Format ID (for built-in formats)
    pub id: Option<u32>,
    /// Format code (for custom formats)
    pub code: Option<String>,
}

impl NumberFormat {
    /// General format
    pub fn general() -> Self {
        Self {
            id: Some(0),
            code: None,
        }
    }

    /// Number format (0.00)
    pub fn number(decimal_places: u8) -> Self {
        let code = if decimal_places == 0 {
            "0".to_string()
        } else {
            format!("0.{}", "0".repeat(decimal_places as usize))
        };
        Self {
            id: None,
            code: Some(code),
        }
    }

    /// Percentage format
    pub fn percentage(decimal_places: u8) -> Self {
        let code = if decimal_places == 0 {
            "0%".to_string()
        } else {
            format!("0.{}%", "0".repeat(decimal_places as usize))
        };
        Self {
            id: None,
            code: Some(code),
        }
    }

    /// Date format
    pub fn date() -> Self {
        Self {
            id: Some(14), // mm-dd-yy
            code: None,
        }
    }

    /// Custom format
    pub fn custom(code: impl Into<String>) -> Self {
        Self {
            id: None,
            code: Some(code.into()),
        }
    }
}

impl Default for NumberFormat {
    fn default() -> Self {
        Self::general()
    }
}

/// Complete cell style.
#[derive(Debug, Clone, PartialEq, Default, Serialize, Deserialize)]
pub struct Style {
    /// Font settings
    pub font: Font,
    /// Fill settings
    pub fill: Fill,
    /// Border settings
    pub border: Border,
    /// Number format
    pub number_format: NumberFormat,
    /// Horizontal alignment
    pub horizontal_alignment: HorizontalAlignment,
    /// Vertical alignment
    pub vertical_alignment: VerticalAlignment,
    /// Text wrap
    pub wrap_text: bool,
    /// Text rotation (degrees, -90 to 90, or 255 for vertical)
    pub text_rotation: Option<i16>,
    /// Indent level
    pub indent: Option<u32>,
    /// Shrink to fit
    pub shrink_to_fit: bool,
}

/// Registry of styles in a workbook.
#[derive(Debug, Clone, Default)]
pub struct StyleRegistry {
    /// Styles by ID
    styles: HashMap<u32, Style>,
    /// Fonts
    fonts: Vec<Font>,
    /// Fills
    fills: Vec<Fill>,
    /// Borders
    borders: Vec<Border>,
    /// Number formats
    number_formats: HashMap<u32, String>,
    /// Next available style ID
    next_id: u32,
}

impl StyleRegistry {
    /// Creates a new empty style registry.
    pub fn new() -> Self {
        Self::default()
    }

    /// Gets a style by ID.
    pub fn get(&self, id: u32) -> Option<&Style> {
        self.styles.get(&id)
    }

    /// Adds a style and returns its ID.
    pub fn add(&mut self, style: Style) -> u32 {
        let id = self.next_id;
        self.styles.insert(id, style);
        self.next_id += 1;
        id
    }

    /// Returns the number of styles.
    pub fn len(&self) -> usize {
        self.styles.len()
    }

    /// Checks if the registry is empty.
    pub fn is_empty(&self) -> bool {
        self.styles.is_empty()
    }

    /// Returns an iterator over all styles.
    pub fn iter(&self) -> impl Iterator<Item = (u32, &Style)> {
        self.styles.iter().map(|(k, v)| (*k, v))
    }

    /// Gets all fonts.
    pub fn fonts(&self) -> &[Font] {
        &self.fonts
    }

    /// Adds a font and returns its index.
    pub fn add_font(&mut self, font: Font) -> usize {
        self.fonts.push(font);
        self.fonts.len() - 1
    }

    /// Gets all fills.
    pub fn fills(&self) -> &[Fill] {
        &self.fills
    }

    /// Adds a fill and returns its index.
    pub fn add_fill(&mut self, fill: Fill) -> usize {
        self.fills.push(fill);
        self.fills.len() - 1
    }

    /// Gets all borders.
    pub fn borders(&self) -> &[Border] {
        &self.borders
    }

    /// Adds a border and returns its index.
    pub fn add_border(&mut self, border: Border) -> usize {
        self.borders.push(border);
        self.borders.len() - 1
    }

    /// Gets a number format by ID.
    pub fn get_number_format(&self, id: u32) -> Option<&str> {
        self.number_formats.get(&id).map(|s| s.as_str())
    }

    /// Returns all number formats.
    pub fn number_formats(&self) -> &HashMap<u32, String> {
        &self.number_formats
    }

    /// Adds a number format and returns its ID.
    pub fn add_number_format(&mut self, id: u32, code: impl Into<String>) {
        self.number_formats.insert(id, code.into());
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_color_rgb() {
        let color = Color::rgb(255, 128, 0);
        assert_eq!(color.to_rgb(), Some((255, 128, 0)));
        assert_eq!(color.to_hex(), Some("FF8000".to_string()));
    }

    #[test]
    fn test_color_rgb_black() {
        let color = Color::rgb(0, 0, 0);
        assert_eq!(color.to_rgb(), Some((0, 0, 0)));
        assert_eq!(color.to_hex(), Some("000000".to_string()));
    }

    #[test]
    fn test_color_rgb_white() {
        let color = Color::rgb(255, 255, 255);
        assert_eq!(color.to_rgb(), Some((255, 255, 255)));
        assert_eq!(color.to_hex(), Some("FFFFFF".to_string()));
    }

    #[test]
    fn test_color_from_hex() {
        assert_eq!(Color::from_hex("FF0000"), Some(Color::Rgb(0xFF0000)));
        assert_eq!(Color::from_hex("#00FF00"), Some(Color::Rgb(0x00FF00)));
        assert_eq!(Color::from_hex("FFFF8000"), Some(Color::Rgb(0xFF8000))); // ARGB
        assert_eq!(Color::from_hex("0000FF"), Some(Color::Rgb(0x0000FF)));
    }

    #[test]
    fn test_color_from_hex_invalid() {
        assert_eq!(Color::from_hex(""), None);
        assert_eq!(Color::from_hex("FF"), None);
        assert_eq!(Color::from_hex("FFFFF"), None);
        assert_eq!(Color::from_hex("GGGGGG"), None); // Invalid hex chars
    }

    #[test]
    fn test_color_theme() {
        let color = Color::Theme(5);
        assert_eq!(color.to_rgb(), None);
        assert_eq!(color.to_hex(), None);
    }

    #[test]
    fn test_color_indexed() {
        let color = Color::Indexed(10);
        assert_eq!(color.to_rgb(), None);
        assert_eq!(color.to_hex(), None);
    }

    #[test]
    fn test_color_auto() {
        let color = Color::Auto;
        assert_eq!(color.to_rgb(), None);
        assert_eq!(color.to_hex(), None);
    }

    #[test]
    fn test_color_default() {
        let color = Color::default();
        assert_eq!(color, Color::Auto);
    }

    #[test]
    fn test_border_all() {
        let border = Border::all(BorderStyle::Thin, Some(Color::rgb(0, 0, 0)));
        assert_eq!(border.left.style, BorderStyle::Thin);
        assert_eq!(border.right.style, BorderStyle::Thin);
        assert_eq!(border.top.style, BorderStyle::Thin);
        assert_eq!(border.bottom.style, BorderStyle::Thin);
        assert_eq!(border.diagonal.style, BorderStyle::None); // diagonal not set by all()
    }

    #[test]
    fn test_border_default() {
        let border = Border::default();
        assert_eq!(border.left.style, BorderStyle::None);
        assert_eq!(border.right.style, BorderStyle::None);
        assert_eq!(border.top.style, BorderStyle::None);
        assert_eq!(border.bottom.style, BorderStyle::None);
        assert!(!border.diagonal_up);
        assert!(!border.diagonal_down);
    }

    #[test]
    fn test_style_registry() {
        let mut registry = StyleRegistry::new();
        assert!(registry.is_empty());

        let style = Style::default();
        let id = registry.add(style.clone());
        assert_eq!(id, 0);
        assert_eq!(registry.len(), 1);

        let retrieved = registry.get(id).unwrap();
        assert_eq!(retrieved, &style);

        // Add another style
        let id2 = registry.add(Style::default());
        assert_eq!(id2, 1);
        assert_eq!(registry.len(), 2);
    }

    #[test]
    fn test_style_registry_get_nonexistent() {
        let registry = StyleRegistry::new();
        assert!(registry.get(999).is_none());
    }

    #[test]
    fn test_style_registry_iter() {
        let mut registry = StyleRegistry::new();
        registry.add(Style::default());
        registry.add(Style::default());
        registry.add(Style::default());

        let items: Vec<_> = registry.iter().collect();
        assert_eq!(items.len(), 3);
    }

    #[test]
    fn test_style_registry_fonts() {
        let mut registry = StyleRegistry::new();

        assert!(registry.fonts().is_empty());

        let font = Font {
            name: Some("Arial".to_string()),
            size: Some(12.0),
            bold: true,
            ..Default::default()
        };

        let idx = registry.add_font(font);
        assert_eq!(idx, 0);
        assert_eq!(registry.fonts().len(), 1);
        assert_eq!(registry.fonts()[0].name, Some("Arial".to_string()));
    }

    #[test]
    fn test_style_registry_fills() {
        let mut registry = StyleRegistry::new();

        assert!(registry.fills().is_empty());

        let fill = Fill {
            pattern: FillPattern::Solid,
            fg_color: Some(Color::rgb(255, 0, 0)),
            bg_color: None,
        };

        let idx = registry.add_fill(fill);
        assert_eq!(idx, 0);
        assert_eq!(registry.fills().len(), 1);
    }

    #[test]
    fn test_style_registry_borders() {
        let mut registry = StyleRegistry::new();

        assert!(registry.borders().is_empty());

        let border = Border::all(BorderStyle::Medium, None);
        let idx = registry.add_border(border);
        assert_eq!(idx, 0);
        assert_eq!(registry.borders().len(), 1);
    }

    #[test]
    fn test_style_registry_number_formats() {
        let mut registry = StyleRegistry::new();

        registry.add_number_format(164, "0.00%");
        registry.add_number_format(165, "#,##0.00");

        assert_eq!(registry.get_number_format(164), Some("0.00%"));
        assert_eq!(registry.get_number_format(165), Some("#,##0.00"));
        assert_eq!(registry.get_number_format(999), None);
    }

    #[test]
    fn test_number_format_general() {
        let fmt = NumberFormat::general();
        assert_eq!(fmt.id, Some(0));
        assert!(fmt.code.is_none());
    }

    #[test]
    fn test_number_format_number() {
        let fmt0 = NumberFormat::number(0);
        assert_eq!(fmt0.code, Some("0".to_string()));

        let fmt2 = NumberFormat::number(2);
        assert_eq!(fmt2.code, Some("0.00".to_string()));

        let fmt4 = NumberFormat::number(4);
        assert_eq!(fmt4.code, Some("0.0000".to_string()));
    }

    #[test]
    fn test_number_format_percentage() {
        let fmt0 = NumberFormat::percentage(0);
        assert_eq!(fmt0.code, Some("0%".to_string()));

        let fmt2 = NumberFormat::percentage(2);
        assert_eq!(fmt2.code, Some("0.00%".to_string()));
    }

    #[test]
    fn test_number_format_date() {
        let fmt = NumberFormat::date();
        assert_eq!(fmt.id, Some(14));
        assert!(fmt.code.is_none());
    }

    #[test]
    fn test_number_format_custom() {
        let fmt = NumberFormat::custom("yyyy-mm-dd hh:mm:ss");
        assert!(fmt.id.is_none());
        assert_eq!(fmt.code, Some("yyyy-mm-dd hh:mm:ss".to_string()));
    }

    #[test]
    fn test_number_format_default() {
        let fmt = NumberFormat::default();
        assert_eq!(fmt, NumberFormat::general());
    }

    #[test]
    fn test_font_default() {
        let font = Font::default();
        assert!(font.name.is_none());
        assert!(font.size.is_none());
        assert!(!font.bold);
        assert!(!font.italic);
        assert!(!font.underline);
        assert!(!font.strikethrough);
        assert!(font.color.is_none());
    }

    #[test]
    fn test_fill_default() {
        let fill = Fill::default();
        assert_eq!(fill.pattern, FillPattern::None);
        assert!(fill.fg_color.is_none());
        assert!(fill.bg_color.is_none());
    }

    #[test]
    fn test_style_default() {
        let style = Style::default();
        assert_eq!(style.horizontal_alignment, HorizontalAlignment::General);
        assert_eq!(style.vertical_alignment, VerticalAlignment::Center);
        assert!(!style.wrap_text);
        assert!(style.text_rotation.is_none());
        assert!(style.indent.is_none());
        assert!(!style.shrink_to_fit);
    }

    #[test]
    fn test_horizontal_alignment_default() {
        assert_eq!(HorizontalAlignment::default(), HorizontalAlignment::General);
    }

    #[test]
    fn test_vertical_alignment_default() {
        assert_eq!(VerticalAlignment::default(), VerticalAlignment::Center);
    }

    #[test]
    fn test_border_style_default() {
        assert_eq!(BorderStyle::default(), BorderStyle::None);
    }

    #[test]
    fn test_fill_pattern_default() {
        assert_eq!(FillPattern::default(), FillPattern::None);
    }

    #[test]
    fn test_border_side_default() {
        let side = BorderSide::default();
        assert_eq!(side.style, BorderStyle::None);
        assert!(side.color.is_none());
    }
}
