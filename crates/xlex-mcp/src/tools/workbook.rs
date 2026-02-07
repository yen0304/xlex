use std::path::PathBuf;

use rmcp::model::{CallToolResult, Content};
use serde::Deserialize;

use crate::error::{session_not_found, xlex_err_to_mcp};
use crate::session::SessionStore;

#[derive(Debug, Deserialize, schemars::JsonSchema)]
pub struct OpenWorkbookParams {
    #[schemars(description = "Path to the xlsx file to open")]
    pub path: String,
}

#[derive(Debug, Deserialize, schemars::JsonSchema)]
pub struct CloseWorkbookParams {
    #[schemars(description = "Session ID returned by open_workbook")]
    pub session_id: String,
    #[schemars(description = "Whether to save before closing (default: false)")]
    pub save: Option<bool>,
}

#[derive(Debug, Deserialize, schemars::JsonSchema)]
pub struct WorkbookInfoParams {
    #[schemars(description = "Session ID returned by open_workbook")]
    pub session_id: String,
}

#[derive(Debug, Deserialize, schemars::JsonSchema)]
pub struct CreateWorkbookParams {
    #[schemars(description = "Path to create the new xlsx file")]
    pub path: String,
    #[schemars(description = "Sheet names to create (default: [\"Sheet1\"])")]
    pub sheets: Option<Vec<String>>,
}

#[derive(Debug, Deserialize, schemars::JsonSchema)]
pub struct SaveWorkbookParams {
    #[schemars(description = "Session ID returned by open_workbook")]
    pub session_id: String,
    #[schemars(description = "Optional new path for save-as")]
    pub path: Option<String>,
}

pub fn open_workbook(store: &SessionStore, params: OpenWorkbookParams) -> CallToolResult {
    let path = PathBuf::from(&params.path);
    match store.open(&path) {
        Ok(session_id) => {
            let info = store
                .with_workbook(&session_id, |wb| {
                    let names = wb.sheet_names();
                    let stats = wb.stats();
                    serde_json::json!({
                        "session_id": session_id,
                        "sheet_count": names.len(),
                        "sheets": names,
                        "file_size": stats.file_size,
                    })
                })
                .unwrap_or_default();
            CallToolResult::success(vec![Content::text(info.to_string())])
        }
        Err(e) => xlex_err_to_mcp(e),
    }
}

pub fn close_workbook(store: &SessionStore, params: CloseWorkbookParams) -> CallToolResult {
    let save = params.save.unwrap_or(false);
    match store.close(&params.session_id, save) {
        Ok(()) => {
            let msg = if save {
                format!("Session {} closed (saved).", params.session_id)
            } else {
                format!("Session {} closed.", params.session_id)
            };
            CallToolResult::success(vec![Content::text(msg)])
        }
        Err(e) => CallToolResult::error(vec![Content::text(e)]),
    }
}

pub fn workbook_info(store: &SessionStore, params: WorkbookInfoParams) -> CallToolResult {
    let Some(result) = store.with_workbook(&params.session_id, |wb| {
        let props = wb.properties();
        let stats = wb.stats();
        serde_json::json!({
            "properties": {
                "title": props.title,
                "creator": props.creator,
                "subject": props.subject,
                "description": props.description,
                "created": props.created,
                "modified": props.modified,
            },
            "stats": {
                "sheet_count": stats.sheet_count,
                "total_cells": stats.total_cells,
                "formula_count": stats.formula_count,
                "style_count": stats.style_count,
                "string_count": stats.string_count,
                "file_size": stats.file_size,
            },
        })
    }) else {
        return session_not_found(&params.session_id);
    };
    CallToolResult::success(vec![Content::text(result.to_string())])
}

pub fn create_workbook(store: &SessionStore, params: CreateWorkbookParams) -> CallToolResult {
    let path = PathBuf::from(&params.path);
    let sheet_names: Vec<&str> = match &params.sheets {
        Some(names) => names.iter().map(|s| s.as_str()).collect(),
        None => vec!["Sheet1"],
    };
    match store.create(&path, &sheet_names) {
        Ok(session_id) => {
            let info = serde_json::json!({
                "session_id": session_id,
                "path": params.path,
                "sheets": sheet_names,
            });
            CallToolResult::success(vec![Content::text(info.to_string())])
        }
        Err(e) => xlex_err_to_mcp(e),
    }
}

pub fn save_workbook(store: &SessionStore, params: SaveWorkbookParams) -> CallToolResult {
    let Some(result) = store.with_workbook_mut(&params.session_id, |wb, original_path| {
        let save_path = params
            .path
            .as_ref()
            .map(PathBuf::from)
            .unwrap_or_else(|| original_path.to_path_buf());
        match wb.save_as(&save_path) {
            Ok(()) => {
                let info = serde_json::json!({
                    "saved_to": save_path.display().to_string(),
                });
                CallToolResult::success(vec![Content::text(info.to_string())])
            }
            Err(e) => xlex_err_to_mcp(e),
        }
    }) else {
        return session_not_found(&params.session_id);
    };
    result
}
