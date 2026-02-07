## Why

AI agents (Claude, GPT, Copilot, etc.) increasingly need to manipulate Excel files directly — reading report data, modifying cells, and converting export formats. xlex-core already provides a complete streaming Excel engine, but it is currently only accessible via the CLI. MCP (Model Context Protocol) is the emerging standard for AI tool interaction. Adding an MCP server allows any MCP-compatible AI agent to invoke xlex capabilities natively, without parsing CLI output.

## What Changes

- Add new workspace crate `crates/xlex-mcp/` as a standalone MCP server binary
- Use `rmcp` crate (official Rust MCP SDK) to implement the MCP protocol
- Wrap xlex-core operations as MCP tools, organized by domain:
  - **Workbook tools**: open, info, stats, create, save
  - **Sheet tools**: list, add, remove, rename
  - **Cell tools**: get, set, clear (with batch support)
  - **Range tools**: get, copy, clear, sort
  - **Export tools**: export to CSV, JSON, Markdown
- Support stdio transport (local use) and Streamable HTTP transport (remote use)
- Provide session management so agents can open a workbook and perform multiple operations
- Add `crates/xlex-mcp` as a workspace member in root Cargo.toml
- Extend CI workflow to cover build, test, and clippy for the new crate

## Capabilities

### New Capabilities

- `mcp-server`: MCP server core framework — transport layer, session management, tool registration, error conversion (XlexError → MCP error)
- `mcp-tools`: MCP tool definitions — wrap xlex-core Workbook/Sheet/Cell/Range/Export operations as MCP tools with input validation schemas and structured response formats

### Modified Capabilities

None — this change does not affect existing xlex-core or xlex-cli behavior.

## Impact

- **New dependencies**: `rmcp` (MCP SDK), `tokio` (already in workspace)
- **Build**: New binary target `xlex-mcp` added to workspace; CI must cover it
- **Release**: New standalone binary; release workflow and install scripts need updates
- **xlex-core**: No modifications required — MCP server depends on existing public API
- **Binary size**: Adds ~5-10MB binary (MCP + tokio runtime included)
