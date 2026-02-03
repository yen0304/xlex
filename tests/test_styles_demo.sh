#!/bin/bash
# XLEX Style Features Demo Script
# Tests: background color, text color, font styles, borders, merged cells

set -e

XLEX="cargo run --quiet --"
OUTPUT="test_styles_demo.xlsx"

echo "=== XLEX Style Features Demo ==="
echo ""

# Clean up old file
rm -f "$OUTPUT"

# Create new workbook
echo "1. Creating workbook..."
$XLEX create "$OUTPUT"

# ===== Background Color Test =====
echo "2. Testing background colors..."
$XLEX cell set "$OUTPUT" Sheet1 A1 "Red Background"
$XLEX cell set "$OUTPUT" Sheet1 A2 "Green Background"
$XLEX cell set "$OUTPUT" Sheet1 A3 "Blue Background"
$XLEX cell set "$OUTPUT" Sheet1 A4 "Yellow Background"
$XLEX cell set "$OUTPUT" Sheet1 A5 "Purple Background"

$XLEX range style "$OUTPUT" Sheet1 A1 --bg-color FF0000
$XLEX range style "$OUTPUT" Sheet1 A2 --bg-color 00FF00
$XLEX range style "$OUTPUT" Sheet1 A3 --bg-color 0000FF
$XLEX range style "$OUTPUT" Sheet1 A4 --bg-color FFFF00
$XLEX range style "$OUTPUT" Sheet1 A5 --bg-color 800080

# ===== Text Color Test =====
echo "3. Testing text colors..."
$XLEX cell set "$OUTPUT" Sheet1 B1 "Red Text"
$XLEX cell set "$OUTPUT" Sheet1 B2 "Green Text"
$XLEX cell set "$OUTPUT" Sheet1 B3 "Blue Text"

$XLEX range style "$OUTPUT" Sheet1 B1 --text-color FF0000
$XLEX range style "$OUTPUT" Sheet1 B2 --text-color 00AA00
$XLEX range style "$OUTPUT" Sheet1 B3 --text-color 0000FF

# ===== Font Style Test =====
echo "4. Testing font styles..."
$XLEX cell set "$OUTPUT" Sheet1 C1 "Bold Text"
$XLEX cell set "$OUTPUT" Sheet1 C2 "Italic Text"
$XLEX cell set "$OUTPUT" Sheet1 C3 "Underline Text"
$XLEX cell set "$OUTPUT" Sheet1 C4 "Bold + Italic"
$XLEX cell set "$OUTPUT" Sheet1 C5 "All Styles"

$XLEX range style "$OUTPUT" Sheet1 C1 --bold
$XLEX range style "$OUTPUT" Sheet1 C2 --italic
$XLEX range style "$OUTPUT" Sheet1 C3 --underline
$XLEX range style "$OUTPUT" Sheet1 C4 --bold --italic
$XLEX range style "$OUTPUT" Sheet1 C5 --bold --italic --underline

# ===== Font Size Test =====
echo "5. Testing font sizes..."
$XLEX cell set "$OUTPUT" Sheet1 D1 "Small (8pt)"
$XLEX cell set "$OUTPUT" Sheet1 D2 "Normal (11pt)"
$XLEX cell set "$OUTPUT" Sheet1 D3 "Large (16pt)"
$XLEX cell set "$OUTPUT" Sheet1 D4 "Extra Large (24pt)"

$XLEX range style "$OUTPUT" Sheet1 D1 --font-size 8
$XLEX range style "$OUTPUT" Sheet1 D2 --font-size 11
$XLEX range style "$OUTPUT" Sheet1 D3 --font-size 16
$XLEX range style "$OUTPUT" Sheet1 D4 --font-size 24

# ===== Border Test =====
echo "6. Testing borders..."
$XLEX cell set "$OUTPUT" Sheet1 E1 "Thin Border"
$XLEX cell set "$OUTPUT" Sheet1 E2 "Medium Border"
$XLEX cell set "$OUTPUT" Sheet1 E3 "Thick Border"

$XLEX range border "$OUTPUT" Sheet1 E1 --style thin
$XLEX range border "$OUTPUT" Sheet1 E2 --style medium
$XLEX range border "$OUTPUT" Sheet1 E3 --style thick

# ===== Merged Cells Test =====
echo "7. Testing merged cells..."
$XLEX cell set "$OUTPUT" Sheet1 F1 "Merged Area F1:G2"
$XLEX range merge "$OUTPUT" Sheet1 F1:G2

$XLEX cell set "$OUTPUT" Sheet1 F4 "Merged Area F4:H4"
$XLEX range merge "$OUTPUT" Sheet1 F4:H4

# ===== Combined Styles Test =====
echo "8. Testing combined styles..."
$XLEX cell set "$OUTPUT" Sheet1 A7 "Combined: Red BG + White Text + Bold"
$XLEX range style "$OUTPUT" Sheet1 A7 --bg-color FF0000 --text-color FFFFFF --bold

$XLEX cell set "$OUTPUT" Sheet1 A8 "Combined: Blue BG + Yellow Text + Italic + Border"
$XLEX range style "$OUTPUT" Sheet1 A8 --bg-color 0000AA --text-color FFFF00 --italic
$XLEX range border "$OUTPUT" Sheet1 A8 --style medium

# ===== Range Style Test =====
echo "9. Testing range styles..."
$XLEX cell set "$OUTPUT" Sheet1 A10 "Range 1"
$XLEX cell set "$OUTPUT" Sheet1 B10 "Range 2"
$XLEX cell set "$OUTPUT" Sheet1 C10 "Range 3"
$XLEX cell set "$OUTPUT" Sheet1 A11 "Range 4"
$XLEX cell set "$OUTPUT" Sheet1 B11 "Range 5"
$XLEX cell set "$OUTPUT" Sheet1 C11 "Range 6"

$XLEX range style "$OUTPUT" Sheet1 A10:C11 --bg-color E0E0FF --bold
$XLEX range border "$OUTPUT" Sheet1 A10:C11 --style thin

# ===== Number Test =====
echo "10. Testing numbers..."
$XLEX cell set "$OUTPUT" Sheet1 H1 "12345.678"
$XLEX cell set "$OUTPUT" Sheet1 H2 "0.5"
$XLEX cell set "$OUTPUT" Sheet1 H3 "1000000"

# ===== Alignment Test =====
echo "11. Testing alignment..."
$XLEX cell set "$OUTPUT" Sheet1 I1 "Left Aligned"
$XLEX cell set "$OUTPUT" Sheet1 I2 "Center Aligned"
$XLEX cell set "$OUTPUT" Sheet1 I3 "Right Aligned"

$XLEX range style "$OUTPUT" Sheet1 I1 --align left
$XLEX range style "$OUTPUT" Sheet1 I2 --align center
$XLEX range style "$OUTPUT" Sheet1 I3 --align right

# ===== Display Results =====
echo ""
echo "=== Test Complete ==="
echo "File created: $(pwd)/$OUTPUT"
echo ""
echo "Style list:"
$XLEX style list "$OUTPUT"
echo ""
echo "Merged cells:"
$XLEX range get "$OUTPUT" Sheet1 F1:H4
echo ""
echo "Open the file with Excel or LibreOffice Calc to verify styles"
