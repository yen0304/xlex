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
#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
pub enum Color {
    /// RGB color (0xRRGGBB)
    Rgb(u32),
    /// Theme color index
    Theme(u32),
    /// Indexed color
    Indexed(u32),
    /// Auto color (black for text, white for background)
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

impl Default for Color {
    fn default() -> Self {
        Self::Auto
    }
}

/// Font style.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
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

impl Default for Font {
    fn default() -> Self {
        Self {
            name: None,
            size: None,
            bold: false,
            italic: false,
            underline: false,
            strikethrough: false,
            color: None,
        }
    }
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
    fn test_color_from_hex() {
        assert_eq!(Color::from_hex("FF0000"), Some(Color::Rgb(0xFF0000)));
        assert_eq!(Color::from_hex("#00FF00"), Some(Color::Rgb(0x00FF00)));
        assert_eq!(Color::from_hex("FFFF8000"), Some(Color::Rgb(0xFF8000))); // ARGB
    }

    #[test]
    fn test_border_all() {
        let border = Border::all(BorderStyle::Thin, Some(Color::rgb(0, 0, 0)));
        assert_eq!(border.left.style, BorderStyle::Thin);
        assert_eq!(border.right.style, BorderStyle::Thin);
        assert_eq!(border.top.style, BorderStyle::Thin);
        assert_eq!(border.bottom.style, BorderStyle::Thin);
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
    }
}
