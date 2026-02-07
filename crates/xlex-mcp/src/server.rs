use rmcp::handler::server::router::tool::ToolRouter;
use rmcp::handler::server::wrapper::Parameters;
use rmcp::model::*;
use rmcp::{tool, tool_handler, tool_router, ServerHandler};

use crate::session::SessionStore;
use crate::tools;

#[derive(Clone)]
pub struct XlexMcpServer {
    store: SessionStore,
    tool_router: ToolRouter<Self>,
}

impl Default for XlexMcpServer {
    fn default() -> Self {
        Self::new()
    }
}

#[tool_router]
impl XlexMcpServer {
    pub fn new() -> Self {
        let store = SessionStore::new();
        Self {
            store,
            tool_router: Self::tool_router(),
        }
    }

    // === Workbook tools ===

    #[tool(
        description = "Open an xlsx file and create a session for subsequent operations. Returns a session_id to use with other tools."
    )]
    fn open_workbook(
        &self,
        Parameters(params): Parameters<tools::workbook::OpenWorkbookParams>,
    ) -> Result<CallToolResult, ErrorData> {
        Ok(tools::workbook::open_workbook(&self.store, params))
    }

    #[tool(description = "Close an open workbook session, optionally saving changes.")]
    fn close_workbook(
        &self,
        Parameters(params): Parameters<tools::workbook::CloseWorkbookParams>,
    ) -> Result<CallToolResult, ErrorData> {
        Ok(tools::workbook::close_workbook(&self.store, params))
    }

    #[tool(
        description = "Get metadata and statistics for an open workbook (properties, sheet count, cell count, etc.)."
    )]
    fn workbook_info(
        &self,
        Parameters(params): Parameters<tools::workbook::WorkbookInfoParams>,
    ) -> Result<CallToolResult, ErrorData> {
        Ok(tools::workbook::workbook_info(&self.store, params))
    }

    #[tool(description = "Create a new empty xlsx workbook file and open it as a session.")]
    fn create_workbook(
        &self,
        Parameters(params): Parameters<tools::workbook::CreateWorkbookParams>,
    ) -> Result<CallToolResult, ErrorData> {
        Ok(tools::workbook::create_workbook(&self.store, params))
    }

    #[tool(description = "Save an open workbook to its original path or a new path (save-as).")]
    fn save_workbook(
        &self,
        Parameters(params): Parameters<tools::workbook::SaveWorkbookParams>,
    ) -> Result<CallToolResult, ErrorData> {
        Ok(tools::workbook::save_workbook(&self.store, params))
    }

    // === Sheet tools ===

    #[tool(
        description = "List all sheets in an open workbook with their name, index, visibility, and cell count."
    )]
    fn list_sheets(
        &self,
        Parameters(params): Parameters<tools::sheet::ListSheetsParams>,
    ) -> Result<CallToolResult, ErrorData> {
        Ok(tools::sheet::list_sheets(&self.store, params))
    }

    #[tool(description = "Add a new sheet to an open workbook.")]
    fn add_sheet(
        &self,
        Parameters(params): Parameters<tools::sheet::AddSheetParams>,
    ) -> Result<CallToolResult, ErrorData> {
        Ok(tools::sheet::add_sheet(&self.store, params))
    }

    #[tool(description = "Remove a sheet from an open workbook by name.")]
    fn remove_sheet(
        &self,
        Parameters(params): Parameters<tools::sheet::RemoveSheetParams>,
    ) -> Result<CallToolResult, ErrorData> {
        Ok(tools::sheet::remove_sheet(&self.store, params))
    }

    #[tool(description = "Rename a sheet in an open workbook.")]
    fn rename_sheet(
        &self,
        Parameters(params): Parameters<tools::sheet::RenameSheetParams>,
    ) -> Result<CallToolResult, ErrorData> {
        Ok(tools::sheet::rename_sheet(&self.store, params))
    }

    // === Cell tools ===

    #[tool(
        description = "Read cell values from a sheet. Supports single cell (cell), range (range in A1:B2 notation), or list of cells (cells)."
    )]
    fn read_cells(
        &self,
        Parameters(params): Parameters<tools::cell::ReadCellsParams>,
    ) -> Result<CallToolResult, ErrorData> {
        Ok(tools::cell::read_cells(&self.store, params))
    }

    #[tool(
        description = "Write values to one or more cells. Each entry specifies a cell reference and value."
    )]
    fn write_cells(
        &self,
        Parameters(params): Parameters<tools::cell::WriteCellsParams>,
    ) -> Result<CallToolResult, ErrorData> {
        Ok(tools::cell::write_cells(&self.store, params))
    }

    #[tool(description = "Clear all cell contents in a range.")]
    fn clear_cells(
        &self,
        Parameters(params): Parameters<tools::cell::ClearCellsParams>,
    ) -> Result<CallToolResult, ErrorData> {
        Ok(tools::cell::clear_cells(&self.store, params))
    }

    // === Row tools ===

    #[tool(description = "Read rows from a sheet with optional start_row and limit.")]
    fn read_rows(
        &self,
        Parameters(params): Parameters<tools::row::ReadRowsParams>,
    ) -> Result<CallToolResult, ErrorData> {
        Ok(tools::row::read_rows(&self.store, params))
    }

    #[tool(description = "Insert empty rows at a specified position, shifting existing rows down.")]
    fn insert_rows(
        &self,
        Parameters(params): Parameters<tools::row::InsertRowsParams>,
    ) -> Result<CallToolResult, ErrorData> {
        Ok(tools::row::insert_rows(&self.store, params))
    }

    #[tool(description = "Delete rows at a specified position, shifting remaining rows up.")]
    fn delete_rows(
        &self,
        Parameters(params): Parameters<tools::row::DeleteRowsParams>,
    ) -> Result<CallToolResult, ErrorData> {
        Ok(tools::row::delete_rows(&self.store, params))
    }

    // === Export tool ===

    #[tool(description = "Export sheet data as CSV, JSON, or Markdown.")]
    fn export_sheet(
        &self,
        Parameters(params): Parameters<tools::export::ExportSheetParams>,
    ) -> Result<CallToolResult, ErrorData> {
        Ok(tools::export::export_sheet(&self.store, params))
    }
}

#[tool_handler]
impl ServerHandler for XlexMcpServer {
    fn get_info(&self) -> ServerInfo {
        ServerInfo {
            protocol_version: ProtocolVersion::V_2025_03_26,
            capabilities: ServerCapabilities::builder().enable_tools().build(),
            server_info: Implementation {
                name: "xlex-mcp".to_string(),
                title: None,
                version: env!("CARGO_PKG_VERSION").to_string(),
                icons: None,
                website_url: None,
            },
            instructions: Some(
                "Excel manipulation MCP server powered by xlex. \
                 Use open_workbook to start a session, then use other tools to read/write cells, \
                 manage sheets, and export data. Close the session with close_workbook when done."
                    .to_string(),
            ),
        }
    }
}
