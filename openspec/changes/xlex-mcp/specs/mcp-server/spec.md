## ADDED Requirements

### Requirement: MCP server binary entry point
The system SHALL provide a standalone binary `xlex-mcp` that starts an MCP server using the `rmcp` crate. The binary SHALL accept a `--transport` flag defaulting to `stdio`.

#### Scenario: Start server with default stdio transport
- **WHEN** user runs `xlex-mcp` with no arguments
- **THEN** the server starts on stdio transport and is ready to accept MCP tool calls

#### Scenario: Start server with explicit transport flag
- **WHEN** user runs `xlex-mcp --transport stdio`
- **THEN** the server starts on stdio transport

#### Scenario: Server reports capabilities on initialize
- **WHEN** an MCP client sends an `initialize` request
- **THEN** the server responds with its name (`xlex-mcp`), version, and the list of available tools

### Requirement: Session-based workbook management
The system SHALL maintain a session store mapping session IDs to open `Workbook` instances. Each `open_workbook` call SHALL create a new session and return a unique session ID. All subsequent tool calls operating on a workbook SHALL require a valid session ID.

#### Scenario: Open workbook creates a session
- **WHEN** an agent calls `open_workbook` with a valid file path
- **THEN** the system opens the file, creates a session with a unique ID, and returns the session ID along with workbook metadata

#### Scenario: Invalid session ID rejected
- **WHEN** an agent calls any workbook operation with a session ID that does not exist in the session store
- **THEN** the system returns an MCP error with a message indicating the session was not found

#### Scenario: Close workbook removes session
- **WHEN** an agent calls `close_workbook` with a valid session ID
- **THEN** the system removes the workbook from the session store and frees its memory

#### Scenario: Multiple concurrent sessions
- **WHEN** an agent opens two different workbooks
- **THEN** both sessions exist independently with separate session IDs, and operations on one do not affect the other

### Requirement: Error conversion from XlexError to MCP error
The system SHALL convert all `XlexError` variants to MCP `CallToolResult` error responses. Each error response SHALL include the XLEX error code, the error message, and the recovery suggestion when available.

#### Scenario: File not found error
- **WHEN** an agent calls `open_workbook` with a path that does not exist
- **THEN** the system returns an MCP error containing `XLEX_E001`, the file path, and a suggestion to check if the file path is correct

#### Scenario: Sheet not found error
- **WHEN** an agent calls `read_cells` referencing a sheet name that does not exist
- **THEN** the system returns an MCP error containing `XLEX_E030`, the sheet name, and a suggestion to use `list_sheets` to see available sheet names

#### Scenario: Error without recovery suggestion
- **WHEN** an xlex-core operation fails with an error that has no recovery suggestion
- **THEN** the system returns an MCP error containing the error code and message without a suggestion line

### Requirement: Workspace crate integration
The `xlex-mcp` crate SHALL be a member of the workspace defined in the root `Cargo.toml`. It SHALL depend on `xlex-core` as a path dependency and use workspace-level dependency versions for shared dependencies (`tokio`, `serde`, `serde_json`).

#### Scenario: Workspace build includes xlex-mcp
- **WHEN** a developer runs `cargo build` from the workspace root
- **THEN** the `xlex-mcp` crate compiles successfully alongside `xlex-core` and `xlex-cli`

#### Scenario: CI pipeline covers xlex-mcp
- **WHEN** CI runs clippy, test, and build jobs
- **THEN** the `xlex-mcp` crate is included in all checks
