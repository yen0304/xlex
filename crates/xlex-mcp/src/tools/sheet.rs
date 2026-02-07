use rmcp::model::{CallToolResult, Content};
use serde::Deserialize;

use crate::error::{session_not_found, xlex_err_to_mcp};
use crate::session::SessionStore;

#[derive(Debug, Deserialize, schemars::JsonSchema)]
pub struct ListSheetsParams {
    #[schemars(description = "Session ID returned by open_workbook")]
    pub session_id: String,
}

#[derive(Debug, Deserialize, schemars::JsonSchema)]
pub struct AddSheetParams {
    #[schemars(description = "Session ID returned by open_workbook")]
    pub session_id: String,
    #[schemars(description = "Name for the new sheet")]
    pub name: String,
}

#[derive(Debug, Deserialize, schemars::JsonSchema)]
pub struct RemoveSheetParams {
    #[schemars(description = "Session ID returned by open_workbook")]
    pub session_id: String,
    #[schemars(description = "Name of the sheet to remove")]
    pub name: String,
}

#[derive(Debug, Deserialize, schemars::JsonSchema)]
pub struct RenameSheetParams {
    #[schemars(description = "Session ID returned by open_workbook")]
    pub session_id: String,
    #[schemars(description = "Current name of the sheet")]
    pub old_name: String,
    #[schemars(description = "New name for the sheet")]
    pub new_name: String,
}

pub fn list_sheets(store: &SessionStore, params: ListSheetsParams) -> CallToolResult {
    let Some(result) = store.with_workbook(&params.session_id, |wb| {
        let mut sheets = Vec::new();
        for (i, name) in wb.sheet_names().iter().enumerate() {
            let cell_count = wb.get_sheet(name).map(|s| s.cell_count()).unwrap_or(0);
            let visibility = wb
                .get_sheet_visibility(name)
                .map(|v| v.to_string())
                .unwrap_or_else(|_| "unknown".to_string());
            sheets.push(serde_json::json!({
                "index": i,
                "name": name,
                "visibility": visibility,
                "cell_count": cell_count,
            }));
        }
        serde_json::json!({ "sheets": sheets })
    }) else {
        return session_not_found(&params.session_id);
    };
    CallToolResult::success(vec![Content::text(result.to_string())])
}

pub fn add_sheet(store: &SessionStore, params: AddSheetParams) -> CallToolResult {
    let Some(result) = store.with_workbook_mut(&params.session_id, |wb, _path| {
        match wb.add_sheet(&params.name) {
            Ok(index) => {
                let info = serde_json::json!({
                    "name": params.name,
                    "index": index,
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

pub fn remove_sheet(store: &SessionStore, params: RemoveSheetParams) -> CallToolResult {
    let Some(result) = store.with_workbook_mut(&params.session_id, |wb, _path| {
        match wb.remove_sheet(&params.name) {
            Ok(()) => CallToolResult::success(vec![Content::text(format!(
                "Sheet '{}' removed.",
                params.name
            ))]),
            Err(e) => xlex_err_to_mcp(e),
        }
    }) else {
        return session_not_found(&params.session_id);
    };
    result
}

pub fn rename_sheet(store: &SessionStore, params: RenameSheetParams) -> CallToolResult {
    let Some(result) = store.with_workbook_mut(&params.session_id, |wb, _path| {
        match wb.rename_sheet(&params.old_name, &params.new_name) {
            Ok(()) => CallToolResult::success(vec![Content::text(format!(
                "Sheet renamed from '{}' to '{}'.",
                params.old_name, params.new_name
            ))]),
            Err(e) => xlex_err_to_mcp(e),
        }
    }) else {
        return session_not_found(&params.session_id);
    };
    result
}
