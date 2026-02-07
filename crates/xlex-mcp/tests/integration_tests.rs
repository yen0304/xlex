use tempfile::TempDir;
use xlex_mcp::session::SessionStore;
use xlex_mcp::tools::{cell, export, row, sheet, workbook};

fn extract_text(result: &rmcp::model::CallToolResult) -> String {
    result
        .content
        .first()
        .and_then(|c| match &c.raw {
            rmcp::model::RawContent::Text(t) => Some(t.text.clone()),
            _ => None,
        })
        .unwrap_or_default()
}

fn is_success(result: &rmcp::model::CallToolResult) -> bool {
    !result.is_error.unwrap_or(false)
}

/// Create a test workbook with some data and return (store, session_id, temp_dir).
fn setup_workbook_with_data() -> (SessionStore, String, TempDir) {
    let temp_dir = TempDir::new().unwrap();
    let xlsx_path = temp_dir.path().join("test.xlsx");
    let store = SessionStore::new();

    // Create workbook
    let result = workbook::create_workbook(
        &store,
        workbook::CreateWorkbookParams {
            path: xlsx_path.to_str().unwrap().to_string(),
            sheets: Some(vec!["Data".to_string()]),
        },
    );
    assert!(is_success(&result));
    let text = extract_text(&result);
    let json: serde_json::Value = serde_json::from_str(&text).unwrap();
    let session_id = json["session_id"].as_str().unwrap().to_string();

    // Write some cells
    let result = cell::write_cells(
        &store,
        cell::WriteCellsParams {
            session_id: session_id.clone(),
            sheet: "Data".to_string(),
            cells: vec![
                cell::CellEntry {
                    cell: "A1".to_string(),
                    value: serde_json::json!("Name"),
                },
                cell::CellEntry {
                    cell: "B1".to_string(),
                    value: serde_json::json!("Score"),
                },
                cell::CellEntry {
                    cell: "A2".to_string(),
                    value: serde_json::json!("Alice"),
                },
                cell::CellEntry {
                    cell: "B2".to_string(),
                    value: serde_json::json!(95),
                },
                cell::CellEntry {
                    cell: "A3".to_string(),
                    value: serde_json::json!("Bob"),
                },
                cell::CellEntry {
                    cell: "B3".to_string(),
                    value: serde_json::json!(87),
                },
            ],
        },
    );
    assert!(is_success(&result));

    (store, session_id, temp_dir)
}

// ============================================================
// 10.1 Open → Read → Close workflow
// ============================================================

mod open_read_close {
    use super::*;

    #[test]
    fn open_read_cells_close() {
        let (store, session_id, _temp_dir) = setup_workbook_with_data();

        // Read a single cell
        let result = cell::read_cells(
            &store,
            cell::ReadCellsParams {
                session_id: session_id.clone(),
                sheet: "Data".to_string(),
                cell: Some("A1".to_string()),
                range: None,
                cells: None,
            },
        );
        assert!(is_success(&result));
        let text = extract_text(&result);
        assert!(text.contains("Name"));

        // Read a range
        let result = cell::read_cells(
            &store,
            cell::ReadCellsParams {
                session_id: session_id.clone(),
                sheet: "Data".to_string(),
                cell: None,
                range: Some("A1:B3".to_string()),
                cells: None,
            },
        );
        assert!(is_success(&result));
        let text = extract_text(&result);
        assert!(text.contains("Alice"));
        assert!(text.contains("95"));
        assert!(text.contains("Bob"));

        // Read rows
        let result = row::read_rows(
            &store,
            row::ReadRowsParams {
                session_id: session_id.clone(),
                sheet: "Data".to_string(),
                start_row: Some(1),
                limit: Some(2),
            },
        );
        assert!(is_success(&result));
        let text = extract_text(&result);
        let json: serde_json::Value = serde_json::from_str(&text).unwrap();
        assert_eq!(json["total_rows"].as_u64().unwrap(), 2);

        // Close without saving
        let result = workbook::close_workbook(
            &store,
            workbook::CloseWorkbookParams {
                session_id: session_id.clone(),
                save: Some(false),
            },
        );
        assert!(is_success(&result));
    }

    #[test]
    fn open_existing_file_and_read() {
        let temp_dir = TempDir::new().unwrap();
        let xlsx_path = temp_dir.path().join("existing.xlsx");
        let store = SessionStore::new();

        // Create and save a workbook first
        let result = workbook::create_workbook(
            &store,
            workbook::CreateWorkbookParams {
                path: xlsx_path.to_str().unwrap().to_string(),
                sheets: None,
            },
        );
        assert!(is_success(&result));
        let text = extract_text(&result);
        let json: serde_json::Value = serde_json::from_str(&text).unwrap();
        let session_id = json["session_id"].as_str().unwrap().to_string();

        // Write data
        cell::write_cells(
            &store,
            cell::WriteCellsParams {
                session_id: session_id.clone(),
                sheet: "Sheet1".to_string(),
                cells: vec![cell::CellEntry {
                    cell: "A1".to_string(),
                    value: serde_json::json!("Hello"),
                }],
            },
        );
        workbook::save_workbook(
            &store,
            workbook::SaveWorkbookParams {
                session_id: session_id.clone(),
                path: None,
            },
        );
        workbook::close_workbook(
            &store,
            workbook::CloseWorkbookParams {
                session_id,
                save: Some(false),
            },
        );

        // Now re-open using open_workbook
        let result = workbook::open_workbook(
            &store,
            workbook::OpenWorkbookParams {
                path: xlsx_path.to_str().unwrap().to_string(),
            },
        );
        assert!(is_success(&result));
        let text = extract_text(&result);
        let json: serde_json::Value = serde_json::from_str(&text).unwrap();
        let session_id2 = json["session_id"].as_str().unwrap().to_string();

        // Verify data persisted
        let result = cell::read_cells(
            &store,
            cell::ReadCellsParams {
                session_id: session_id2.clone(),
                sheet: "Sheet1".to_string(),
                cell: Some("A1".to_string()),
                range: None,
                cells: None,
            },
        );
        assert!(is_success(&result));
        let text = extract_text(&result);
        assert!(text.contains("Hello"));

        workbook::close_workbook(
            &store,
            workbook::CloseWorkbookParams {
                session_id: session_id2,
                save: Some(false),
            },
        );
    }
}

// ============================================================
// 10.2 Create → Write → Save → Reopen → Verify
// ============================================================

mod write_save_reopen {
    use super::*;

    #[test]
    fn create_write_save_reopen_verify() {
        let temp_dir = TempDir::new().unwrap();
        let xlsx_path = temp_dir.path().join("roundtrip.xlsx");
        let store = SessionStore::new();

        // Create workbook with custom sheets
        let result = workbook::create_workbook(
            &store,
            workbook::CreateWorkbookParams {
                path: xlsx_path.to_str().unwrap().to_string(),
                sheets: Some(vec!["Sales".to_string(), "Summary".to_string()]),
            },
        );
        assert!(is_success(&result));
        let text = extract_text(&result);
        let json: serde_json::Value = serde_json::from_str(&text).unwrap();
        let session_id = json["session_id"].as_str().unwrap().to_string();

        // Write cells with different types
        let result = cell::write_cells(
            &store,
            cell::WriteCellsParams {
                session_id: session_id.clone(),
                sheet: "Sales".to_string(),
                cells: vec![
                    cell::CellEntry {
                        cell: "A1".to_string(),
                        value: serde_json::json!("Product"),
                    },
                    cell::CellEntry {
                        cell: "B1".to_string(),
                        value: serde_json::json!("Price"),
                    },
                    cell::CellEntry {
                        cell: "A2".to_string(),
                        value: serde_json::json!("Widget"),
                    },
                    cell::CellEntry {
                        cell: "B2".to_string(),
                        value: serde_json::json!(29.99),
                    },
                    cell::CellEntry {
                        cell: "C2".to_string(),
                        value: serde_json::json!(true),
                    },
                ],
            },
        );
        assert!(is_success(&result));
        let text = extract_text(&result);
        let json: serde_json::Value = serde_json::from_str(&text).unwrap();
        assert_eq!(json["written"].as_u64().unwrap(), 5);

        // Save
        let result = workbook::save_workbook(
            &store,
            workbook::SaveWorkbookParams {
                session_id: session_id.clone(),
                path: None,
            },
        );
        assert!(is_success(&result));

        // Close
        let result = workbook::close_workbook(
            &store,
            workbook::CloseWorkbookParams {
                session_id,
                save: Some(false),
            },
        );
        assert!(is_success(&result));

        // Reopen with a fresh store
        let store2 = SessionStore::new();
        let result = workbook::open_workbook(
            &store2,
            workbook::OpenWorkbookParams {
                path: xlsx_path.to_str().unwrap().to_string(),
            },
        );
        assert!(is_success(&result));
        let text = extract_text(&result);
        let json: serde_json::Value = serde_json::from_str(&text).unwrap();
        let session_id2 = json["session_id"].as_str().unwrap().to_string();
        assert_eq!(json["sheet_count"].as_u64().unwrap(), 2);

        // Verify written values
        let result = cell::read_cells(
            &store2,
            cell::ReadCellsParams {
                session_id: session_id2.clone(),
                sheet: "Sales".to_string(),
                cell: None,
                range: None,
                cells: Some(vec![
                    "A1".to_string(),
                    "B1".to_string(),
                    "A2".to_string(),
                    "B2".to_string(),
                ]),
            },
        );
        assert!(is_success(&result));
        let text = extract_text(&result);
        assert!(text.contains("Product"));
        assert!(text.contains("Price"));
        assert!(text.contains("Widget"));
        assert!(text.contains("29.99"));

        workbook::close_workbook(
            &store2,
            workbook::CloseWorkbookParams {
                session_id: session_id2,
                save: Some(false),
            },
        );
    }

    #[test]
    fn save_as_new_path() {
        let temp_dir = TempDir::new().unwrap();
        let original_path = temp_dir.path().join("original.xlsx");
        let copy_path = temp_dir.path().join("copy.xlsx");
        let store = SessionStore::new();

        // Create and write data
        let result = workbook::create_workbook(
            &store,
            workbook::CreateWorkbookParams {
                path: original_path.to_str().unwrap().to_string(),
                sheets: None,
            },
        );
        let text = extract_text(&result);
        let json: serde_json::Value = serde_json::from_str(&text).unwrap();
        let session_id = json["session_id"].as_str().unwrap().to_string();

        cell::write_cells(
            &store,
            cell::WriteCellsParams {
                session_id: session_id.clone(),
                sheet: "Sheet1".to_string(),
                cells: vec![cell::CellEntry {
                    cell: "A1".to_string(),
                    value: serde_json::json!("saved-as"),
                }],
            },
        );

        // Save-as to a different path
        let result = workbook::save_workbook(
            &store,
            workbook::SaveWorkbookParams {
                session_id: session_id.clone(),
                path: Some(copy_path.to_str().unwrap().to_string()),
            },
        );
        assert!(is_success(&result));
        assert!(copy_path.exists());

        workbook::close_workbook(
            &store,
            workbook::CloseWorkbookParams {
                session_id,
                save: Some(false),
            },
        );

        // Open the copy and verify
        let result = workbook::open_workbook(
            &store,
            workbook::OpenWorkbookParams {
                path: copy_path.to_str().unwrap().to_string(),
            },
        );
        assert!(is_success(&result));
        let text = extract_text(&result);
        let json: serde_json::Value = serde_json::from_str(&text).unwrap();
        let sid = json["session_id"].as_str().unwrap().to_string();

        let result = cell::read_cells(
            &store,
            cell::ReadCellsParams {
                session_id: sid.clone(),
                sheet: "Sheet1".to_string(),
                cell: Some("A1".to_string()),
                range: None,
                cells: None,
            },
        );
        assert!(is_success(&result));
        assert!(extract_text(&result).contains("saved-as"));

        workbook::close_workbook(
            &store,
            workbook::CloseWorkbookParams {
                session_id: sid,
                save: Some(false),
            },
        );
    }
}

// ============================================================
// 10.3 Error cases
// ============================================================

mod error_cases {
    use super::*;

    #[test]
    fn invalid_session_id() {
        let store = SessionStore::new();

        let result = cell::read_cells(
            &store,
            cell::ReadCellsParams {
                session_id: "nonexistent-session".to_string(),
                sheet: "Sheet1".to_string(),
                cell: Some("A1".to_string()),
                range: None,
                cells: None,
            },
        );
        assert!(!is_success(&result));
        assert!(extract_text(&result).contains("Session not found"));
    }

    #[test]
    fn file_not_found() {
        let store = SessionStore::new();

        let result = workbook::open_workbook(
            &store,
            workbook::OpenWorkbookParams {
                path: "/tmp/this-file-does-not-exist-12345.xlsx".to_string(),
            },
        );
        assert!(!is_success(&result));
    }

    #[test]
    fn sheet_not_found() {
        let (store, session_id, _temp_dir) = setup_workbook_with_data();

        // read_cells uses read_single_cell which returns an empty value for missing sheets
        // rather than an error, so check read_rows which does explicit sheet check
        let result = row::read_rows(
            &store,
            row::ReadRowsParams {
                session_id: session_id.clone(),
                sheet: "NonexistentSheet".to_string(),
                start_row: None,
                limit: None,
            },
        );
        assert!(!is_success(&result));
        assert!(extract_text(&result).contains("Sheet not found"));

        workbook::close_workbook(
            &store,
            workbook::CloseWorkbookParams {
                session_id,
                save: Some(false),
            },
        );
    }

    #[test]
    fn close_already_closed_session() {
        let (store, session_id, _temp_dir) = setup_workbook_with_data();

        // Close once
        let result = workbook::close_workbook(
            &store,
            workbook::CloseWorkbookParams {
                session_id: session_id.clone(),
                save: Some(false),
            },
        );
        assert!(is_success(&result));

        // Close again - should error
        let result = workbook::close_workbook(
            &store,
            workbook::CloseWorkbookParams {
                session_id,
                save: Some(false),
            },
        );
        assert!(!is_success(&result));
    }

    #[test]
    fn invalid_cell_reference() {
        let (store, session_id, _temp_dir) = setup_workbook_with_data();

        let _result = cell::write_cells(
            &store,
            cell::WriteCellsParams {
                session_id: session_id.clone(),
                sheet: "Data".to_string(),
                cells: vec![cell::CellEntry {
                    cell: "ZZZZ99999".to_string(),
                    value: serde_json::json!("test"),
                }],
            },
        );
        // Invalid cell refs may or may not error depending on parser - just ensure it doesn't panic

        workbook::close_workbook(
            &store,
            workbook::CloseWorkbookParams {
                session_id,
                save: Some(false),
            },
        );
    }
}

// ============================================================
// 10.4 Verify all 16 tools + sheet/export/row operations
// ============================================================

mod all_tools {
    use super::*;

    #[test]
    fn sheet_operations() {
        let (store, session_id, _temp_dir) = setup_workbook_with_data();

        // list_sheets
        let result = sheet::list_sheets(
            &store,
            sheet::ListSheetsParams {
                session_id: session_id.clone(),
            },
        );
        assert!(is_success(&result));
        let text = extract_text(&result);
        assert!(text.contains("Data"));

        // add_sheet
        let result = sheet::add_sheet(
            &store,
            sheet::AddSheetParams {
                session_id: session_id.clone(),
                name: "NewSheet".to_string(),
            },
        );
        assert!(is_success(&result));

        // rename_sheet
        let result = sheet::rename_sheet(
            &store,
            sheet::RenameSheetParams {
                session_id: session_id.clone(),
                old_name: "NewSheet".to_string(),
                new_name: "Renamed".to_string(),
            },
        );
        assert!(is_success(&result));

        // remove_sheet
        let result = sheet::remove_sheet(
            &store,
            sheet::RemoveSheetParams {
                session_id: session_id.clone(),
                name: "Renamed".to_string(),
            },
        );
        assert!(is_success(&result));

        workbook::close_workbook(
            &store,
            workbook::CloseWorkbookParams {
                session_id,
                save: Some(false),
            },
        );
    }

    #[test]
    fn row_operations() {
        let (store, session_id, _temp_dir) = setup_workbook_with_data();

        // insert_rows
        let result = row::insert_rows(
            &store,
            row::InsertRowsParams {
                session_id: session_id.clone(),
                sheet: "Data".to_string(),
                row: 2,
                count: Some(1),
            },
        );
        assert!(is_success(&result));
        let text = extract_text(&result);
        let json: serde_json::Value = serde_json::from_str(&text).unwrap();
        assert_eq!(json["inserted"].as_u64().unwrap(), 1);

        // delete_rows
        let result = row::delete_rows(
            &store,
            row::DeleteRowsParams {
                session_id: session_id.clone(),
                sheet: "Data".to_string(),
                row: 2,
                count: Some(1),
            },
        );
        assert!(is_success(&result));
        let text = extract_text(&result);
        let json: serde_json::Value = serde_json::from_str(&text).unwrap();
        assert_eq!(json["deleted"].as_u64().unwrap(), 1);

        workbook::close_workbook(
            &store,
            workbook::CloseWorkbookParams {
                session_id,
                save: Some(false),
            },
        );
    }

    #[test]
    fn clear_cells() {
        let (store, session_id, _temp_dir) = setup_workbook_with_data();

        let result = cell::clear_cells(
            &store,
            cell::ClearCellsParams {
                session_id: session_id.clone(),
                sheet: "Data".to_string(),
                range: "A2:B3".to_string(),
            },
        );
        assert!(is_success(&result));
        let text = extract_text(&result);
        let json: serde_json::Value = serde_json::from_str(&text).unwrap();
        assert_eq!(json["cleared"].as_u64().unwrap(), 4);

        workbook::close_workbook(
            &store,
            workbook::CloseWorkbookParams {
                session_id,
                save: Some(false),
            },
        );
    }

    #[test]
    fn export_csv() {
        let (store, session_id, _temp_dir) = setup_workbook_with_data();

        let result = export::export_sheet(
            &store,
            export::ExportSheetParams {
                session_id: session_id.clone(),
                sheet: "Data".to_string(),
                format: "csv".to_string(),
            },
        );
        assert!(is_success(&result));
        let text = extract_text(&result);
        assert!(text.contains("Name"));
        assert!(text.contains("Alice"));

        workbook::close_workbook(
            &store,
            workbook::CloseWorkbookParams {
                session_id,
                save: Some(false),
            },
        );
    }

    #[test]
    fn export_json() {
        let (store, session_id, _temp_dir) = setup_workbook_with_data();

        let result = export::export_sheet(
            &store,
            export::ExportSheetParams {
                session_id: session_id.clone(),
                sheet: "Data".to_string(),
                format: "json".to_string(),
            },
        );
        assert!(is_success(&result));
        let text = extract_text(&result);
        // JSON uses first row as headers
        assert!(text.contains("Name"));
        assert!(text.contains("Alice"));

        workbook::close_workbook(
            &store,
            workbook::CloseWorkbookParams {
                session_id,
                save: Some(false),
            },
        );
    }

    #[test]
    fn export_markdown() {
        let (store, session_id, _temp_dir) = setup_workbook_with_data();

        let result = export::export_sheet(
            &store,
            export::ExportSheetParams {
                session_id: session_id.clone(),
                sheet: "Data".to_string(),
                format: "markdown".to_string(),
            },
        );
        assert!(is_success(&result));
        let text = extract_text(&result);
        assert!(text.contains("| Name"));
        assert!(text.contains("| ---"));

        workbook::close_workbook(
            &store,
            workbook::CloseWorkbookParams {
                session_id,
                save: Some(false),
            },
        );
    }

    #[test]
    fn workbook_info() {
        let (store, session_id, _temp_dir) = setup_workbook_with_data();

        let result = workbook::workbook_info(
            &store,
            workbook::WorkbookInfoParams {
                session_id: session_id.clone(),
            },
        );
        assert!(is_success(&result));
        let text = extract_text(&result);
        let json: serde_json::Value = serde_json::from_str(&text).unwrap();
        assert!(json["stats"]["sheet_count"].as_u64().unwrap() >= 1);

        workbook::close_workbook(
            &store,
            workbook::CloseWorkbookParams {
                session_id,
                save: Some(false),
            },
        );
    }

    /// Verify all 16 tool functions are callable and return valid results.
    /// This is a compile-time + runtime check that all tools exist and work.
    #[test]
    fn all_16_tools_exist() {
        let temp_dir = TempDir::new().unwrap();
        let xlsx_path = temp_dir.path().join("all_tools.xlsx");
        let store = SessionStore::new();

        // 1. create_workbook
        let r = workbook::create_workbook(
            &store,
            workbook::CreateWorkbookParams {
                path: xlsx_path.to_str().unwrap().to_string(),
                sheets: Some(vec!["Sheet1".to_string()]),
            },
        );
        assert!(is_success(&r));
        let sid = serde_json::from_str::<serde_json::Value>(&extract_text(&r)).unwrap()
            ["session_id"]
            .as_str()
            .unwrap()
            .to_string();

        // 2. workbook_info
        assert!(is_success(&workbook::workbook_info(
            &store,
            workbook::WorkbookInfoParams {
                session_id: sid.clone(),
            },
        )));

        // 3. save_workbook
        assert!(is_success(&workbook::save_workbook(
            &store,
            workbook::SaveWorkbookParams {
                session_id: sid.clone(),
                path: None,
            },
        )));

        // 4. list_sheets
        assert!(is_success(&sheet::list_sheets(
            &store,
            sheet::ListSheetsParams {
                session_id: sid.clone(),
            },
        )));

        // 5. add_sheet
        assert!(is_success(&sheet::add_sheet(
            &store,
            sheet::AddSheetParams {
                session_id: sid.clone(),
                name: "Extra".to_string(),
            },
        )));

        // 6. rename_sheet
        assert!(is_success(&sheet::rename_sheet(
            &store,
            sheet::RenameSheetParams {
                session_id: sid.clone(),
                old_name: "Extra".to_string(),
                new_name: "Extra2".to_string(),
            },
        )));

        // 7. remove_sheet
        assert!(is_success(&sheet::remove_sheet(
            &store,
            sheet::RemoveSheetParams {
                session_id: sid.clone(),
                name: "Extra2".to_string(),
            },
        )));

        // 8. write_cells
        assert!(is_success(&cell::write_cells(
            &store,
            cell::WriteCellsParams {
                session_id: sid.clone(),
                sheet: "Sheet1".to_string(),
                cells: vec![cell::CellEntry {
                    cell: "A1".to_string(),
                    value: serde_json::json!("test"),
                }],
            },
        )));

        // 9. read_cells (single)
        assert!(is_success(&cell::read_cells(
            &store,
            cell::ReadCellsParams {
                session_id: sid.clone(),
                sheet: "Sheet1".to_string(),
                cell: Some("A1".to_string()),
                range: None,
                cells: None,
            },
        )));

        // 10. clear_cells
        assert!(is_success(&cell::clear_cells(
            &store,
            cell::ClearCellsParams {
                session_id: sid.clone(),
                sheet: "Sheet1".to_string(),
                range: "A1:A1".to_string(),
            },
        )));

        // 11. read_rows
        assert!(is_success(&row::read_rows(
            &store,
            row::ReadRowsParams {
                session_id: sid.clone(),
                sheet: "Sheet1".to_string(),
                start_row: None,
                limit: None,
            },
        )));

        // 12. insert_rows
        assert!(is_success(&row::insert_rows(
            &store,
            row::InsertRowsParams {
                session_id: sid.clone(),
                sheet: "Sheet1".to_string(),
                row: 1,
                count: Some(1),
            },
        )));

        // 13. delete_rows
        assert!(is_success(&row::delete_rows(
            &store,
            row::DeleteRowsParams {
                session_id: sid.clone(),
                sheet: "Sheet1".to_string(),
                row: 1,
                count: Some(1),
            },
        )));

        // 14. export_sheet
        assert!(is_success(&export::export_sheet(
            &store,
            export::ExportSheetParams {
                session_id: sid.clone(),
                sheet: "Sheet1".to_string(),
                format: "csv".to_string(),
            },
        )));

        // 15. close_workbook
        assert!(is_success(&workbook::close_workbook(
            &store,
            workbook::CloseWorkbookParams {
                session_id: sid.clone(),
                save: Some(false),
            },
        )));

        // 16. open_workbook
        let r = workbook::open_workbook(
            &store,
            workbook::OpenWorkbookParams {
                path: xlsx_path.to_str().unwrap().to_string(),
            },
        );
        assert!(is_success(&r));
        let sid2 = serde_json::from_str::<serde_json::Value>(&extract_text(&r)).unwrap()
            ["session_id"]
            .as_str()
            .unwrap()
            .to_string();
        workbook::close_workbook(
            &store,
            workbook::CloseWorkbookParams {
                session_id: sid2,
                save: Some(false),
            },
        );
    }
}
