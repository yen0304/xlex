## Context

xlex is a workspace with two crates: `xlex-core` (streaming Excel engine) and `xlex-cli` (CLI interface). The goal is to add a third crate, `xlex-mcp`, that exposes xlex-core's capabilities as an MCP server. This allows AI agents (Claude Code, Cursor, etc.) to manipulate Excel files directly through the MCP protocol without parsing CLI output.

The `xlex-core` public API provides:
- `Workbook` — full in-memory workbook with read/write operations
- `LazyWorkbook` — streaming read-only access for large files
- `Sheet`, `Cell`, `CellRef`, `CellValue`, `Range` — domain types
- `XlexError` — structured errors with error codes and recovery suggestions

Key constraint: `Workbook` holds a `ZipArchive<Cursor<Vec<u8>>>` internally. It is not `Send` or `Sync` by default. The MCP server must handle concurrency carefully.

## Goals / Non-Goals

**Goals:**
- Expose xlex-core operations as MCP tools usable by any MCP-compatible AI agent
- Support stdio transport for local use (Claude Code, Cursor, IDE integrations)
- Provide session-based workbook management (open once, operate many times)
- Return structured responses that AI agents can reason about
- Map xlex-core errors to actionable MCP error responses

**Non-Goals:**
- Streamable HTTP transport (defer to a future change — adds complexity with auth, CORS, etc.)
- Formula evaluation (xlex-core does not support this)
- VBA/macro execution (security boundary)
- Real-time collaboration or multi-user access
- File watching or auto-reload

## Decisions

### 1. MCP SDK: `rmcp`

**Decision**: Use the `rmcp` crate (official Rust MCP SDK, v0.14+).

**Rationale**: It is the official Rust implementation maintained under `modelcontextprotocol/rust-sdk`. It provides `#[tool]` and `#[tool_router]` macros for declarative tool registration, built-in stdio transport, and tokio-based async runtime (already in the workspace).

**Alternatives considered**:
- `rust-mcp-sdk`: Less mature, smaller community
- Custom implementation: Unnecessary effort when an official SDK exists
- TypeScript/Python MCP server calling xlex CLI as subprocess: Adds language boundary, loses type safety, slower

### 2. Session Management: `HashMap<SessionId, Workbook>` behind `Arc<Mutex<>>`

**Decision**: Maintain a session store that maps session IDs to open `Workbook` instances. Each `open_workbook` call creates a session, returns a session ID. Subsequent tool calls reference the session ID.

```
SessionStore {
    workbooks: Arc<Mutex<HashMap<String, Workbook>>>
}
```

**Rationale**: AI agents typically work with a file across multiple tool calls (open → read → modify → save). Without sessions, each tool call would need to re-open and re-parse the file. Sessions also match the existing CLI session mode pattern.

The `Mutex` is required because `Workbook` is not `Send`/`Sync` (contains `ZipArchive`). Since MCP tool calls are sequential per-client (the protocol does not require concurrent tool execution), mutex contention is minimal.

**Alternatives considered**:
- Stateless (re-open file each call): Simple but slow for multi-step workflows
- One workbook per server: Too limiting — agents may need to work with multiple files
- `RwLock`: Not beneficial since most operations mutate the workbook

### 3. Tool Granularity: Domain-Oriented, Medium Granularity

**Decision**: Organize tools by domain with medium granularity. Each tool maps to one xlex-core operation. Use clear `domain_action` naming.

| Tool Name | Description |
|-----------|-------------|
| `open_workbook` | Open an xlsx file, return session ID and workbook info |
| `close_workbook` | Close a session, optionally save |
| `workbook_info` | Get workbook metadata, properties, stats |
| `list_sheets` | List all sheets with metadata |
| `add_sheet` | Add a new sheet |
| `remove_sheet` | Remove a sheet by name |
| `rename_sheet` | Rename a sheet |
| `read_cells` | Read one or more cells (supports single cell, range, or list) |
| `write_cells` | Write one or more cells (supports batch) |
| `clear_cells` | Clear cells in a range |
| `read_rows` | Read rows from a sheet (with optional range/limit) |
| `insert_rows` | Insert empty rows |
| `delete_rows` | Delete rows |
| `save_workbook` | Save workbook (to original or new path) |
| `export_sheet` | Export sheet data as CSV, JSON, or Markdown |
| `create_workbook` | Create a new empty workbook |

**Rationale**: 16 tools is manageable for AI agents without being too coarse. Each tool has a clear single responsibility. The `read_cells` tool supports both single-cell and range reads to reduce round-trips.

**Alternatives considered**:
- Fine-grained (30+ tools, e.g., separate `get_cell`/`get_range`/`get_row`): Too many tools for agents to reason about
- Coarse-grained (5 tools with complex params): Hard for agents to discover capabilities; complex parameter schemas
- CRUD-style (`workbook_create`/`workbook_read`/etc.): Less intuitive naming

### 4. Response Format: Structured JSON with Markdown Summary

**Decision**: Each tool returns `Content::text()` with a structured format containing:
1. A human-readable markdown summary (for display)
2. Embedded JSON data blocks (for programmatic use)

For tabular data (cells, rows), return markdown tables for readability with raw JSON available via an optional `format` parameter.

**Rationale**: AI agents consume text responses. Markdown is readable and parseable. Including structured data in the response lets agents extract specific values without regex parsing.

### 5. Error Mapping: XlexError → MCP Error with Recovery Suggestions

**Decision**: Map `XlexError` to MCP `CallToolResult::error()` responses, including:
- The XLEX error code (e.g., `XLEX_E030`)
- The error message
- The recovery suggestion from `XlexError::recovery_suggestion()`

```rust
fn xlex_err_to_mcp(err: XlexError) -> CallToolResult {
    let msg = format!("[{}] {}", err.code(), err);
    let suggestion = err.recovery_suggestion().unwrap_or_default();
    CallToolResult::error(vec![Content::text(format!("{}\n\nSuggestion: {}", msg, suggestion))])
}
```

**Rationale**: AI agents benefit from actionable error messages. xlex-core already provides recovery suggestions — we expose them directly.

### 6. Crate Structure

```
crates/xlex-mcp/
├── Cargo.toml
└── src/
    ├── main.rs          # Entry point, transport setup, CLI args
    ├── server.rs        # MCP server struct, tool_router impl
    ├── session.rs       # SessionStore, workbook lifecycle
    ├── tools/
    │   ├── mod.rs       # Tool registration
    │   ├── workbook.rs  # open, close, info, create, save
    │   ├── sheet.rs     # list, add, remove, rename
    │   ├── cell.rs      # read_cells, write_cells, clear_cells
    │   ├── row.rs       # read_rows, insert_rows, delete_rows
    │   └── export.rs    # export_sheet
    └── error.rs         # XlexError → MCP error conversion
```

### 7. Transport: stdio Only (for Now)

**Decision**: Ship with stdio transport only. Streamable HTTP is a future enhancement.

**Rationale**: The primary use case is local AI agent integration (Claude Code, Cursor). All major MCP clients support stdio. HTTP transport adds authentication, CORS, binding, and security concerns that are orthogonal to the core MCP tool implementation.

## Risks / Trade-offs

**[Mutex contention on SessionStore]** → Mitigated by the sequential nature of MCP tool calls per client. If future multi-client support is needed, can switch to per-session locks.

**[Memory usage with multiple open workbooks]** → Mitigated by `close_workbook` tool and documenting that agents should close files when done. Can add session timeout/eviction in the future.

**[rmcp API stability (pre-1.0)]** → Mitigated by pinning to a specific version range. The `#[tool]` macro API is stable across recent releases.

**[Workbook is not Send/Sync]** → Mitigated by wrapping in `Mutex`. Since tool calls are sequential, this does not affect performance. If xlex-core later makes Workbook thread-safe, the Mutex can be removed.

**[Large file handling]** → For read-only large file operations, `LazyWorkbook` with streaming could be used. Initial implementation uses `Workbook` for simplicity. Can add a `stream_rows` tool backed by `LazyWorkbook` later.
