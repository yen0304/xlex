use rmcp::model::{CallToolResult, Content};
use serde::Deserialize;
use xlex_core::CellRef;

use crate::error::{session_not_found, xlex_err_to_mcp};
use crate::session::SessionStore;

#[derive(Debug, Deserialize, schemars::JsonSchema)]
pub struct ExportSheetParams {
    #[schemars(description = "Session ID returned by open_workbook")]
    pub session_id: String,
    #[schemars(description = "Sheet name")]
    pub sheet: String,
    #[schemars(description = "Export format: csv, json, or markdown")]
    pub format: String,
}

pub fn export_sheet(store: &SessionStore, params: ExportSheetParams) -> CallToolResult {
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
        if max_col == 0 || max_row == 0 {
            return CallToolResult::success(vec![Content::text("(empty sheet)")]);
        }

        // Collect all data as strings
        let mut rows: Vec<Vec<String>> = Vec::new();
        for row in 1..=max_row {
            let mut row_data = Vec::new();
            for col in 1..=max_col {
                let cell_ref = CellRef::new(col, row);
                let value = sheet.get_value(&cell_ref);
                row_data.push(value.to_display_string());
            }
            rows.push(row_data);
        }

        match params.format.as_str() {
            "csv" => {
                let csv = rows
                    .iter()
                    .map(|row| {
                        row.iter()
                            .map(|cell| {
                                if cell.contains(',') || cell.contains('"') || cell.contains('\n') {
                                    format!("\"{}\"", cell.replace('"', "\"\""))
                                } else {
                                    cell.clone()
                                }
                            })
                            .collect::<Vec<_>>()
                            .join(",")
                    })
                    .collect::<Vec<_>>()
                    .join("\n");
                CallToolResult::success(vec![Content::text(csv)])
            }
            "json" => {
                if rows.is_empty() {
                    return CallToolResult::success(vec![Content::text("[]")]);
                }
                let headers = &rows[0];
                let data_rows: Vec<serde_json::Value> = rows[1..]
                    .iter()
                    .map(|row| {
                        let mut obj = serde_json::Map::new();
                        for (i, cell) in row.iter().enumerate() {
                            let key = headers
                                .get(i)
                                .cloned()
                                .unwrap_or_else(|| format!("col_{}", i + 1));
                            obj.insert(key, serde_json::json!(cell));
                        }
                        serde_json::Value::Object(obj)
                    })
                    .collect();
                CallToolResult::success(vec![Content::text(
                    serde_json::to_string_pretty(&data_rows).unwrap_or_else(|_| "[]".to_string()),
                )])
            }
            "markdown" => {
                if rows.is_empty() {
                    return CallToolResult::success(vec![Content::text("(empty)")]);
                }
                let mut md = String::new();
                // Header row
                md.push_str("| ");
                md.push_str(&rows[0].join(" | "));
                md.push_str(" |\n");
                // Separator
                md.push_str("| ");
                md.push_str(
                    &rows[0]
                        .iter()
                        .map(|_| "---")
                        .collect::<Vec<_>>()
                        .join(" | "),
                );
                md.push_str(" |\n");
                // Data rows
                for row in &rows[1..] {
                    md.push_str("| ");
                    md.push_str(&row.join(" | "));
                    md.push_str(" |\n");
                }
                CallToolResult::success(vec![Content::text(md)])
            }
            other => CallToolResult::error(vec![Content::text(format!(
                "Unsupported format: {other}. Use csv, json, or markdown."
            ))]),
        }
    }) else {
        return session_not_found(&params.session_id);
    };
    result
}
