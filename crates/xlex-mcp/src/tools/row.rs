use rmcp::model::{CallToolResult, Content};
use serde::Deserialize;
use xlex_core::CellRef;

use crate::error::{session_not_found, xlex_err_to_mcp};
use crate::session::SessionStore;
use crate::tools::cell::cell_value_to_json;

#[derive(Debug, Deserialize, schemars::JsonSchema)]
pub struct ReadRowsParams {
    #[schemars(description = "Session ID returned by open_workbook")]
    pub session_id: String,
    #[schemars(description = "Sheet name")]
    pub sheet: String,
    #[schemars(description = "Start row number (1-indexed, default: 1)")]
    pub start_row: Option<u32>,
    #[schemars(description = "Maximum number of rows to read (default: all)")]
    pub limit: Option<u32>,
}

#[derive(Debug, Deserialize, schemars::JsonSchema)]
pub struct InsertRowsParams {
    #[schemars(description = "Session ID returned by open_workbook")]
    pub session_id: String,
    #[schemars(description = "Sheet name")]
    pub sheet: String,
    #[schemars(description = "Row position to insert at (1-indexed)")]
    pub row: u32,
    #[schemars(description = "Number of rows to insert (default: 1)")]
    pub count: Option<u32>,
}

#[derive(Debug, Deserialize, schemars::JsonSchema)]
pub struct DeleteRowsParams {
    #[schemars(description = "Session ID returned by open_workbook")]
    pub session_id: String,
    #[schemars(description = "Sheet name")]
    pub sheet: String,
    #[schemars(description = "Row position to delete at (1-indexed)")]
    pub row: u32,
    #[schemars(description = "Number of rows to delete (default: 1)")]
    pub count: Option<u32>,
}

pub fn read_rows(store: &SessionStore, params: ReadRowsParams) -> CallToolResult {
    let Some(result) = store.with_workbook(&params.session_id, |wb| {
        let sheet = match wb.get_sheet(&params.sheet) {
            Some(s) => s,
            None => {
                return xlex_err_to_mcp(xlex_core::XlexError::SheetNotFound {
                    name: params.sheet.clone(),
                })
            }
        };
        let (max_col, max_row) = sheet.dimensions();
        let start = params.start_row.unwrap_or(1);
        let end = match params.limit {
            Some(limit) => (start + limit - 1).min(max_row),
            None => max_row,
        };

        let mut rows = Vec::new();
        for row in start..=end {
            let mut cells = Vec::new();
            for col in 1..=max_col {
                let cell_ref = CellRef::new(col, row);
                let value = sheet.get_value(&cell_ref);
                if !matches!(value, xlex_core::CellValue::Empty) {
                    let mut cell_json = cell_value_to_json(&value);
                    if let serde_json::Value::Object(ref mut map) = cell_json {
                        map.insert("cell".to_string(), serde_json::json!(cell_ref.to_string()));
                    }
                    cells.push(cell_json);
                }
            }
            if !cells.is_empty() {
                rows.push(serde_json::json!({"row": row, "cells": cells}));
            }
        }

        CallToolResult::success(vec![Content::text(
            serde_json::json!({"rows": rows, "total_rows": rows.len()}).to_string(),
        )])
    }) else {
        return session_not_found(&params.session_id);
    };
    result
}

pub fn insert_rows(store: &SessionStore, params: InsertRowsParams) -> CallToolResult {
    let Some(result) = store.with_workbook_mut(&params.session_id, |wb, _path| {
        let sheet = match wb.get_sheet_mut(&params.sheet) {
            Some(s) => s,
            None => {
                return xlex_err_to_mcp(xlex_core::XlexError::SheetNotFound {
                    name: params.sheet.clone(),
                })
            }
        };
        let count = params.count.unwrap_or(1);
        sheet.insert_rows(params.row, count);
        CallToolResult::success(vec![Content::text(
            serde_json::json!({
                "inserted": count,
                "at_row": params.row,
                "sheet": params.sheet,
            })
            .to_string(),
        )])
    }) else {
        return session_not_found(&params.session_id);
    };
    result
}

pub fn delete_rows(store: &SessionStore, params: DeleteRowsParams) -> CallToolResult {
    let Some(result) = store.with_workbook_mut(&params.session_id, |wb, _path| {
        let sheet = match wb.get_sheet_mut(&params.sheet) {
            Some(s) => s,
            None => {
                return xlex_err_to_mcp(xlex_core::XlexError::SheetNotFound {
                    name: params.sheet.clone(),
                })
            }
        };
        let count = params.count.unwrap_or(1);
        sheet.delete_rows(params.row, count);
        CallToolResult::success(vec![Content::text(
            serde_json::json!({
                "deleted": count,
                "at_row": params.row,
                "sheet": params.sheet,
            })
            .to_string(),
        )])
    }) else {
        return session_not_found(&params.session_id);
    };
    result
}
