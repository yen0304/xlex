use rmcp::model::{CallToolResult, Content};
use xlex_core::XlexError;

/// Convert an `XlexError` into an MCP error response.
pub fn xlex_err_to_mcp(err: XlexError) -> CallToolResult {
    let code = err.code();
    let suggestion = err.recovery_suggestion().unwrap_or_default();
    let message = if suggestion.is_empty() {
        format!("[{code}] {err}")
    } else {
        format!("[{code}] {err}\n\nSuggestion: {suggestion}")
    };
    CallToolResult::error(vec![Content::text(message)])
}

/// Convert a session-not-found error into an MCP error response.
pub fn session_not_found(session_id: &str) -> CallToolResult {
    CallToolResult::error(vec![Content::text(format!(
        "Session not found: {session_id}\n\nSuggestion: Use open_workbook to create a session first, or check the session ID."
    ))])
}
