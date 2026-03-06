## 1. Project Scaffolding

- [x] 1.1 Create `crates/xlex-mcp/` directory with `Cargo.toml` (depend on `xlex-core` path, `rmcp`, `tokio`, `serde`, `serde_json`, `clap`)
- [x] 1.2 Add `"crates/xlex-mcp"` to workspace members in root `Cargo.toml`
- [x] 1.3 Create `src/main.rs` with clap CLI args (`--transport stdio`) and tokio async main
- [x] 1.4 Verify `cargo build -p xlex-mcp` compiles successfully

## 2. Error Handling

- [x] 2.1 Create `src/error.rs` with `xlex_err_to_mcp()` function that converts `XlexError` to `CallToolResult::error()` including error code, message, and recovery suggestion

## 3. Session Management

- [x] 3.1 Create `src/session.rs` with `SessionStore` struct (`Arc<Mutex<HashMap<String, Workbook>>>`)
- [x] 3.2 Implement `SessionStore::open()` — open workbook, generate UUID session ID, insert into store, return ID
- [x] 3.3 Implement `SessionStore::close()` — remove workbook from store, optionally save before removal
- [x] 3.4 Implement `SessionStore::with_workbook()` — borrow a workbook by session ID for read operations
- [x] 3.5 Implement `SessionStore::with_workbook_mut()` — mutably borrow a workbook by session ID for write operations

## 4. MCP Server Core

- [x] 4.1 Create `src/server.rs` with `XlexMcpServer` struct holding `SessionStore`
- [x] 4.2 Implement `ServerHandler` trait for `XlexMcpServer` (server name, version, capabilities)
- [x] 4.3 Wire up `main.rs` to create `XlexMcpServer` and serve on stdio transport via `rmcp`

## 5. Workbook Tools

- [x] 5.1 Create `src/tools/mod.rs` and `src/tools/workbook.rs`
- [x] 5.2 Implement `open_workbook` tool — accept `path`, call `SessionStore::open()`, return session ID + workbook metadata
- [x] 5.3 Implement `close_workbook` tool — accept `session_id` and optional `save`, call `SessionStore::close()`
- [x] 5.4 Implement `workbook_info` tool — accept `session_id`, return document properties and stats
- [x] 5.5 Implement `create_workbook` tool — accept `path` and optional `sheets`, create file, open as session
- [x] 5.6 Implement `save_workbook` tool — accept `session_id` and optional `path`, save to original or new path

## 6. Sheet Tools

- [x] 6.1 Create `src/tools/sheet.rs`
- [x] 6.2 Implement `list_sheets` tool — return name, index, visibility, cell count for each sheet
- [x] 6.3 Implement `add_sheet` tool — accept `name`, add sheet, return index
- [x] 6.4 Implement `remove_sheet` tool — accept `name`, remove sheet
- [x] 6.5 Implement `rename_sheet` tool — accept `old_name` and `new_name`

## 7. Cell Tools

- [x] 7.1 Create `src/tools/cell.rs`
- [x] 7.2 Implement `read_cells` tool — support single `cell`, `range` (A1:B2), and `cells` (list) modes
- [x] 7.3 Implement `write_cells` tool — accept list of `{cell, value}` entries, support string/number/boolean types
- [x] 7.4 Implement `clear_cells` tool — accept `range`, clear all cells in range

## 8. Row Tools

- [x] 8.1 Create `src/tools/row.rs`
- [x] 8.2 Implement `read_rows` tool — accept `start_row` and `limit`, return rows with cell values
- [x] 8.3 Implement `insert_rows` tool — accept `row` position and `count`
- [x] 8.4 Implement `delete_rows` tool — accept `row` position and `count`

## 9. Export Tool

- [x] 9.1 Create `src/tools/export.rs`
- [x] 9.2 Implement `export_sheet` tool — accept `format` (csv/json/markdown), return exported content as text

## 10. Integration & Testing

- [x] 10.1 Add integration test that starts the server, opens a workbook, reads cells, and closes the session
- [x] 10.2 Add integration test for write workflow: create workbook → write cells → save → reopen → verify values
- [x] 10.3 Add integration test for error cases: invalid session ID, file not found, sheet not found
- [x] 10.4 Verify all 16 tools appear in `initialize` response tool list

## 11. CI & Release

- [x] 11.1 Update `.github/workflows/ci.yml` to include `xlex-mcp` in build, test, and clippy jobs
- [x] 11.2 Update `.github/workflows/release.yml` to build and publish `xlex-mcp` binary
