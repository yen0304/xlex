# XLEX

<div align="center">

**CLI-first streaming Excel manipulation tool for developers**

[![Crates.io](https://img.shields.io/crates/v/xlex-cli.svg)](https://crates.io/crates/xlex-cli)
[![License](https://img.shields.io/badge/license-MIT-blue.svg)](LICENSE)
[![CI](https://github.com/yen0304/xlex/workflows/CI/badge.svg)](https://github.com/yen0304/xlex/actions)

[Installation](#installation) •
[Quick Start](#quick-start) •
[Commands](commands/index.md) •
[Guides](guides/pipelines.md)

</div>

---

## What is XLEX?

XLEX is a high-performance CLI tool for working with Excel files (.xlsx). It's designed for developers, DevOps engineers, and anyone who needs to automate Excel operations in scripts and pipelines.

### Key Features

- 🚀 **Streaming Architecture** - Handle files up to 200MB without memory exhaustion
- 🔧 **CLI-First Design** - Built for Unix pipelines and automation
- 📊 **Full Excel Support** - Cells, formulas, styles, charts, and more
- 🎨 **Template Engine** - Generate reports with Handlebars-like syntax
- 💡 **Developer Friendly** - JSON/CSV output, meaningful error codes

## Quick Example

```bash
# Create a workbook and add data
xlex create report.xlsx
xlex cell set report.xlsx Sheet1 A1 "Sales Report"
xlex row append report.xlsx Sheet1 "Q1,10000,5000"
xlex row append report.xlsx Sheet1 "Q2,12000,6000"

# Export to CSV
xlex export csv report.xlsx sales.csv

# Apply a template
xlex template apply template.xlsx data.json output.xlsx
```

## Why XLEX?

| Feature | XLEX | Python Libraries | GUI Tools |
|---------|------|------------------|-----------|
| Streaming (large files) | ✅ | ❌ | ❌ |
| CLI/Pipeline friendly | ✅ | ⚠️ | ❌ |
| No dependencies | ✅ | ❌ | ❌ |
| Cross-platform | ✅ | ✅ | ⚠️ |
| Template engine | ✅ | ⚠️ | ❌ |

## Installation

=== "Cargo"

    ```bash
    cargo install xlex-cli
    ```

=== "Homebrew"

    ```bash
    brew install user/tap/xlex
    ```

=== "npm/npx"

    ```bash
    npx xlex-cli --help
    # or install globally
    npm install -g xlex-cli
    ```

=== "Shell Script"

    ```bash
    curl -fsSL https://raw.githubusercontent.com/yen0304/xlex/main/install.sh | bash
    ```

## Getting Help

- 📖 Read the [documentation](https://yen0304.github.io/xlex)
- 💬 [GitHub Discussions](https://github.com/yen0304/xlex/discussions)
- 🐛 [Issue Tracker](https://github.com/yen0304/xlex/issues)

## License

XLEX is licensed under the [MIT License](https://github.com/yen0304/xlex/blob/main/LICENSE).
