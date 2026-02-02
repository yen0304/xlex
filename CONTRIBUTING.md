# Contributing to xlex

Thank you for your interest in contributing to xlex! This document provides guidelines and instructions for contributing.

## Code of Conduct

Please be respectful and constructive in all interactions.

## Getting Started

### Prerequisites

- Rust 1.75 or later
- Git

### Setup

```bash
# Clone the repository
git clone https://github.com/yourusername/xlex.git
cd xlex

# Build the project
cargo build

# Run tests
cargo test

# Run the CLI
cargo run -- info sample.xlsx
```

## Development Workflow

### Branch Naming

- `feature/<name>` - New features
- `fix/<name>` - Bug fixes
- `docs/<name>` - Documentation changes
- `refactor/<name>` - Code refactoring

### Commit Messages

Follow [Conventional Commits](https://www.conventionalcommits.org/):

```
type(scope): description

[optional body]

[optional footer]
```

Types: `feat`, `fix`, `docs`, `style`, `refactor`, `test`, `chore`

Examples:
```
feat(cell): add support for date values
fix(parser): handle empty shared strings
docs(readme): update installation instructions
```

## Code Guidelines

### Style

- Follow Rust standard style (enforced by `cargo fmt`)
- Use meaningful variable and function names
- Document public APIs with doc comments

### Linting

```bash
# Format code
cargo fmt

# Run clippy
cargo clippy -- -D warnings
```

### Testing

```bash
# Run all tests
cargo test

# Run tests with output
cargo test -- --nocapture

# Run specific test
cargo test test_name
```

### Documentation

```bash
# Build documentation
cargo doc --no-deps --open
```

## Pull Request Process

1. Create a feature branch from `main`
2. Make your changes
3. Add tests for new functionality
4. Update documentation if needed
5. Run `cargo fmt` and `cargo clippy`
6. Push and create a Pull Request

### PR Checklist

- [ ] Code follows style guidelines
- [ ] Tests pass locally
- [ ] Documentation updated
- [ ] Commit messages follow conventions
- [ ] PR description explains changes

## Architecture

### Crate Structure

```
xlex/
├── crates/
│   ├── xlex-core/        # Core library
│   │   ├── src/
│   │   │   ├── lib.rs
│   │   │   ├── cell.rs       # Cell types
│   │   │   ├── range.rs      # Range handling
│   │   │   ├── sheet.rs      # Sheet operations
│   │   │   ├── style.rs      # Styling
│   │   │   ├── workbook.rs   # Main workbook type
│   │   │   ├── error.rs      # Error types
│   │   │   ├── parser/       # XLSX parsing
│   │   │   └── writer/       # XLSX writing
│   │   └── Cargo.toml
│   └── xlex-cli/         # CLI binary
│       ├── src/
│       │   ├── main.rs
│       │   └── commands/     # CLI commands
│       └── Cargo.toml
└── Cargo.toml            # Workspace root
```

### Key Design Decisions

1. **Streaming XML Parsing**: Uses SAX-style parsing (quick-xml) to handle large files
2. **Lazy Loading**: SharedStrings loaded on-demand with LRU cache
3. **Copy-on-Write**: Modifications use COW pattern for efficiency
4. **Atomic Saves**: Uses temp files + rename for safe writes

## Adding a New Command

1. Create a new file in `crates/xlex-cli/src/commands/`
2. Define arguments using clap derive macros
3. Implement the command function
4. Add to `mod.rs` and `Commands` enum
5. Add to `Cli::run()` match

Example:

```rust
// commands/example.rs
use clap::{Parser, Subcommand};
use anyhow::Result;

#[derive(Parser)]
pub struct ExampleArgs {
    #[command(subcommand)]
    pub command: ExampleCommand,
}

#[derive(Subcommand)]
pub enum ExampleCommand {
    /// Do something
    Do {
        file: std::path::PathBuf,
    },
}

pub fn run(args: &ExampleArgs, global: &super::GlobalOptions) -> Result<()> {
    match &args.command {
        ExampleCommand::Do { file } => {
            // Implementation
            Ok(())
        }
    }
}
```

## Reporting Issues

### Bug Reports

Include:
- xlex version (`xlex --version`)
- Operating system
- Steps to reproduce
- Expected vs actual behavior
- Sample file if possible

### Feature Requests

Include:
- Use case description
- Proposed CLI interface
- Examples of expected behavior

## Questions?

Open a [discussion](https://github.com/yourusername/xlex/discussions) for questions.
