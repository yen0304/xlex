# Test Fixtures

This directory contains sample xlsx files for testing XLEX CLI commands.

## Files

- `simple.xlsx` - A simple workbook with basic data
- `multi-sheet.xlsx` - Workbook with multiple sheets
- `formulas.xlsx` - Workbook with various formulas
- `styles.xlsx` - Workbook with cell styles and formatting
- `template.xlsx` - Sample template with placeholders
- `large.xlsx` - Large file for performance testing (generated)

## Generating Fixtures

Run the fixture generator to create or regenerate test files:

```bash
cargo run --bin xlex -- create tests/fixtures/simple.xlsx
```

Or use the provided script:

```bash
./tests/generate_fixtures.sh
```

## Notes

- Some fixtures are generated programmatically to ensure consistency
- The `large.xlsx` file is not committed to git due to size
- Use `tests/fixtures/.gitignore` to exclude generated large files
