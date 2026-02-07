use rmcp::model::{CallToolResult, Content};
use serde::Deserialize;
use xlex_core::{CellRef, CellValue};

use crate::error::{session_not_found, xlex_err_to_mcp};
use crate::session::SessionStore;

#[derive(Debug, Deserialize, schemars::JsonSchema)]
pub struct ReadCellsParams {
    #[schemars(description = "Session ID returned by open_workbook")]
    pub session_id: String,
    #[schemars(description = "Sheet name")]
    pub sheet: String,
    #[schemars(description = "Single cell reference in A1 notation (e.g. \"A1\")")]
    pub cell: Option<String>,
    #[schemars(description = "Range in A1:B2 notation (e.g. \"A1:C3\")")]
    pub range: Option<String>,
    #[schemars(description = "List of cell references (e.g. [\"A1\", \"B5\", \"D10\"])")]
    pub cells: Option<Vec<String>>,
}

#[derive(Debug, Deserialize, schemars::JsonSchema)]
pub struct CellEntry {
    #[schemars(description = "Cell reference in A1 notation")]
    pub cell: String,
    #[schemars(description = "Value to write (string, number, or boolean)")]
    pub value: serde_json::Value,
}

#[derive(Debug, Deserialize, schemars::JsonSchema)]
pub struct WriteCellsParams {
    #[schemars(description = "Session ID returned by open_workbook")]
    pub session_id: String,
    #[schemars(description = "Sheet name")]
    pub sheet: String,
    #[schemars(description = "List of cells to write")]
    pub cells: Vec<CellEntry>,
}

#[derive(Debug, Deserialize, schemars::JsonSchema)]
pub struct ClearCellsParams {
    #[schemars(description = "Session ID returned by open_workbook")]
    pub session_id: String,
    #[schemars(description = "Sheet name")]
    pub sheet: String,
    #[schemars(description = "Range to clear in A1:B2 notation")]
    pub range: String,
}

pub fn cell_value_to_json(value: &CellValue) -> serde_json::Value {
    match value {
        CellValue::Empty => serde_json::Value::Null,
        CellValue::String(s) => serde_json::json!({"type": "string", "value": s}),
        CellValue::Number(n) => serde_json::json!({"type": "number", "value": n}),
        CellValue::Boolean(b) => serde_json::json!({"type": "boolean", "value": b}),
        CellValue::Error(e) => serde_json::json!({"type": "error", "value": e.to_string()}),
        CellValue::Formula {
            formula,
            cached_result,
        } => {
            let cached = cached_result
                .as_ref()
                .map(|v| v.to_display_string())
                .unwrap_or_default();
            serde_json::json!({"type": "formula", "formula": formula, "cached_value": cached})
        }
        CellValue::DateTime(n) => serde_json::json!({"type": "datetime", "value": n}),
    }
}

fn json_to_cell_value(value: &serde_json::Value) -> CellValue {
    match value {
        serde_json::Value::Null => CellValue::Empty,
        serde_json::Value::Bool(b) => CellValue::Boolean(*b),
        serde_json::Value::Number(n) => CellValue::Number(n.as_f64().unwrap_or(0.0)),
        serde_json::Value::String(s) => {
            if s.starts_with('=') {
                CellValue::Formula {
                    formula: s.clone(),
                    cached_result: None,
                }
            } else {
                CellValue::String(s.clone())
            }
        }
        _ => CellValue::String(value.to_string()),
    }
}

fn read_single_cell(wb: &xlex_core::Workbook, sheet: &str, cell_str: &str) -> serde_json::Value {
    match CellRef::parse(cell_str) {
        Ok(cell_ref) => match wb.get_cell(sheet, &cell_ref) {
            Ok(value) => {
                let mut result = cell_value_to_json(&value);
                if let serde_json::Value::Object(ref mut map) = result {
                    map.insert(
                        "cell".to_string(),
                        serde_json::json!(cell_str.to_uppercase()),
                    );
                } else {
                    result = serde_json::json!({"cell": cell_str.to_uppercase(), "type": "empty", "value": null});
                }
                result
            }
            Err(_) => {
                serde_json::json!({"cell": cell_str.to_uppercase(), "type": "empty", "value": null})
            }
        },
        Err(_) => serde_json::json!({"error": format!("Invalid cell reference: {cell_str}")}),
    }
}

pub fn read_cells(store: &SessionStore, params: ReadCellsParams) -> CallToolResult {
    let Some(result) = store.with_workbook(&params.session_id, |wb| {
        let mut results = Vec::new();

        // Single cell mode
        if let Some(ref cell) = params.cell {
            results.push(read_single_cell(wb, &params.sheet, cell));
        }

        // Range mode
        if let Some(ref range_str) = params.range {
            match xlex_core::Range::parse(range_str) {
                Ok(range) => {
                    for row in range.start.row..=range.end.row {
                        for col in range.start.col..=range.end.col {
                            let cell_ref = CellRef::new(col, row);
                            let cell_str = cell_ref.to_string();
                            results.push(read_single_cell(wb, &params.sheet, &cell_str));
                        }
                    }
                }
                Err(e) => return xlex_err_to_mcp(e),
            }
        }

        // List mode
        if let Some(ref cells) = params.cells {
            for cell_str in cells {
                results.push(read_single_cell(wb, &params.sheet, cell_str));
            }
        }

        if results.is_empty() {
            return CallToolResult::error(vec![Content::text(
                "No cells specified. Provide one of: cell, range, or cells parameter.",
            )]);
        }

        CallToolResult::success(vec![Content::text(
            serde_json::json!({"cells": results}).to_string(),
        )])
    }) else {
        return session_not_found(&params.session_id);
    };
    result
}

pub fn write_cells(store: &SessionStore, params: WriteCellsParams) -> CallToolResult {
    let Some(result) = store.with_workbook_mut(&params.session_id, |wb, _path| {
        let mut written = 0;
        for entry in &params.cells {
            let cell_ref = match CellRef::parse(&entry.cell) {
                Ok(r) => r,
                Err(e) => return xlex_err_to_mcp(e),
            };
            let value = json_to_cell_value(&entry.value);
            if let Err(e) = wb.set_cell(&params.sheet, cell_ref, value) {
                return xlex_err_to_mcp(e);
            }
            written += 1;
        }
        CallToolResult::success(vec![Content::text(
            serde_json::json!({"written": written, "sheet": params.sheet}).to_string(),
        )])
    }) else {
        return session_not_found(&params.session_id);
    };
    result
}

pub fn clear_cells(store: &SessionStore, params: ClearCellsParams) -> CallToolResult {
    let Some(result) = store.with_workbook_mut(&params.session_id, |wb, _path| {
        let range = match xlex_core::Range::parse(&params.range) {
            Ok(r) => r,
            Err(e) => return xlex_err_to_mcp(e),
        };
        let mut cleared = 0;
        for row in range.start.row..=range.end.row {
            for col in range.start.col..=range.end.col {
                let cell_ref = CellRef::new(col, row);
                if let Err(e) = wb.clear_cell(&params.sheet, &cell_ref) {
                    return xlex_err_to_mcp(e);
                }
                cleared += 1;
            }
        }
        CallToolResult::success(vec![Content::text(
            serde_json::json!({"cleared": cleared, "range": params.range}).to_string(),
        )])
    }) else {
        return session_not_found(&params.session_id);
    };
    result
}
