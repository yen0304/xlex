# xlex-mcp

An [MCP (Model Context Protocol)](https://modelcontextprotocol.io/) server for Excel file manipulation, powered by [xlex](https://github.com/yen0304/xlex).

This server allows AI assistants like Claude to read, write, and manipulate `.xlsx` files through a session-based API with 16 tools for workbook, sheet, cell, row, and export operations.

## Installation

### From Source

```bash
git clone https://github.com/yen0304/xlex.git
cd xlex
cargo install --path crates/xlex-mcp
```

### From Release Binary

Download the `xlex-mcp` binary from the [releases page](https://github.com/yen0304/xlex/releases).

## Configuration

### Claude Desktop

Add to your `claude_desktop_config.json`:

```json
{
  "mcpServers": {
    "xlex": {
      "command": "xlex-mcp",
      "args": []
    }
  }
}
```

If the binary is not in your `PATH`, use the full path:

```json
{
  "mcpServers": {
    "xlex": {
      "command": "/path/to/xlex-mcp",
      "args": []
    }
  }
}
```

### Claude Code

Add to your project's `.mcp.json`:

```json
{
  "mcpServers": {
    "xlex": {
      "command": "xlex-mcp",
      "args": []
    }
  }
}
```

### Generic stdio

The server communicates over stdio using the MCP protocol:

```bash
xlex-mcp --transport stdio
```

## Available Tools

### Workbook (5 tools)

| Tool | Description |
|------|-------------|
| `open_workbook` | Open an xlsx file and create a session. Returns a `session_id` for subsequent operations. |
| `close_workbook` | Close a session, optionally saving changes. |
| `workbook_info` | Get metadata and statistics (properties, sheet count, cell count, etc.). |
| `create_workbook` | Create a new empty xlsx file and open it as a session. |
| `save_workbook` | Save to the original path or a new path (save-as). |

### Sheet (4 tools)

| Tool | Description |
|------|-------------|
| `list_sheets` | List all sheets with name, index, visibility, and cell count. |
| `add_sheet` | Add a new sheet. |
| `remove_sheet` | Remove a sheet by name. |
| `rename_sheet` | Rename a sheet. |

### Cell (3 tools)

| Tool | Description |
|------|-------------|
| `read_cells` | Read cell values by single cell, range (`A1:B2`), or list of cells. |
| `write_cells` | Write values to one or more cells. Supports strings, numbers, booleans, and formulas (`=SUM(A1:A10)`). |
| `clear_cells` | Clear all cell contents in a range. |

### Row (3 tools)

| Tool | Description |
|------|-------------|
| `read_rows` | Read rows with optional `start_row` and `limit`. |
| `insert_rows` | Insert empty rows at a position, shifting existing rows down. |
| `delete_rows` | Delete rows at a position, shifting remaining rows up. |

### Export (1 tool)

| Tool | Description |
|------|-------------|
| `export_sheet` | Export sheet data as CSV, JSON, or Markdown. |

## Usage Workflow

A typical workflow follows this pattern:

```
open_workbook → read/write operations → save_workbook → close_workbook
```

**Example — read data from a spreadsheet:**

1. `open_workbook` with `path: "report.xlsx"` → get `session_id`
2. `list_sheets` → see available sheets
3. `read_cells` with `range: "A1:D10"` → read the data
4. `close_workbook` → end the session

**Example — create and populate a new spreadsheet:**

1. `create_workbook` with `path: "output.xlsx"` → get `session_id`
2. `write_cells` with cell entries → populate data
3. `save_workbook` → persist to disk
4. `close_workbook` → end the session

## Architecture

The server uses a **session-based** model:

- Each `open_workbook` or `create_workbook` call creates an isolated session with a unique `session_id`.
- Multiple workbooks can be open simultaneously in separate sessions.
- Workbook data is held in memory for the duration of the session.
- Changes are only written to disk when `save_workbook` is called or `close_workbook` is called with `save: true`.
- Sessions are cleaned up when `close_workbook` is called.

## License

MIT License — see [LICENSE](../../LICENSE) for details.
