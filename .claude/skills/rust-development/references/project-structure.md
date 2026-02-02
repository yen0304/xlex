# Rust Project Structure

## Binary Project

```
my-app/
├── Cargo.toml
├── Cargo.lock
├── .gitignore
├── README.md
├── src/
│   ├── main.rs           # Entry point
│   ├── lib.rs            # Library root (optional)
│   ├── config.rs         # Configuration module
│   ├── error.rs          # Error types
│   └── utils/
│       ├── mod.rs
│       └── helpers.rs
├── tests/
│   └── integration_test.rs
├── benches/
│   └── benchmarks.rs
└── examples/
    └── basic_usage.rs
```

## Library Project

```
my-lib/
├── Cargo.toml
├── src/
│   ├── lib.rs            # Public API exports
│   ├── core/
│   │   ├── mod.rs
│   │   ├── types.rs
│   │   └── traits.rs
│   ├── utils/
│   │   ├── mod.rs
│   │   └── helpers.rs
│   └── error.rs
├── tests/
│   ├── common/
│   │   └── mod.rs        # Shared test utilities
│   └── api_tests.rs
├── benches/
│   └── performance.rs
└── examples/
    ├── simple.rs
    └── advanced.rs
```

## Workspace (Monorepo)

```
my-workspace/
├── Cargo.toml            # Workspace root
├── crates/
│   ├── core/
│   │   ├── Cargo.toml
│   │   └── src/
│   │       └── lib.rs
│   ├── api/
│   │   ├── Cargo.toml
│   │   └── src/
│   │       └── lib.rs
│   └── cli/
│       ├── Cargo.toml
│       └── src/
│           └── main.rs
├── tests/
│   └── integration/
└── docs/
```

### Workspace Cargo.toml
```toml
[workspace]
resolver = "2"
members = [
    "crates/*",
]

[workspace.package]
version = "0.1.0"
edition = "2021"
license = "MIT"
repository = "https://github.com/user/repo"

[workspace.dependencies]
tokio = { version = "1", features = ["full"] }
serde = { version = "1", features = ["derive"] }
anyhow = "1"
```

### Member Cargo.toml
```toml
[package]
name = "my-core"
version.workspace = true
edition.workspace = true

[dependencies]
tokio.workspace = true
serde.workspace = true
```

## Web Service Project

```
my-service/
├── Cargo.toml
├── src/
│   ├── main.rs
│   ├── lib.rs
│   ├── config.rs
│   ├── error.rs
│   ├── routes/
│   │   ├── mod.rs
│   │   ├── health.rs
│   │   └── api/
│   │       ├── mod.rs
│   │       ├── users.rs
│   │       └── items.rs
│   ├── handlers/
│   │   ├── mod.rs
│   │   └── user_handler.rs
│   ├── models/
│   │   ├── mod.rs
│   │   └── user.rs
│   ├── services/
│   │   ├── mod.rs
│   │   └── user_service.rs
│   └── db/
│       ├── mod.rs
│       └── migrations/
├── migrations/
│   └── 001_initial.sql
├── tests/
│   └── api_tests.rs
└── docker/
    ├── Dockerfile
    └── docker-compose.yml
```

## CLI Application

```
my-cli/
├── Cargo.toml
├── src/
│   ├── main.rs
│   ├── cli.rs            # Clap definitions
│   ├── commands/
│   │   ├── mod.rs
│   │   ├── init.rs
│   │   ├── build.rs
│   │   └── run.rs
│   ├── config.rs
│   └── error.rs
├── tests/
│   └── cli_tests.rs
└── completions/
    ├── bash/
    └── zsh/
```

## Module Organization

### mod.rs Pattern (Traditional)
```
src/
├── lib.rs
└── utils/
    ├── mod.rs            # pub mod helpers;
    └── helpers.rs
```

### File-as-Module Pattern (Modern, Rust 2018+)
```
src/
├── lib.rs
├── utils.rs              # Module root
└── utils/
    └── helpers.rs        # Submodule
```

### lib.rs Example
```rust
// Re-export public API
pub mod config;
pub mod error;

mod internal;  // Private module

pub use config::Config;
pub use error::{Error, Result};
```

### mod.rs Example
```rust
// src/utils/mod.rs
mod helpers;
mod formatters;

pub use helpers::*;
pub use formatters::format_output;
```

## Configuration Files

### Essential Files
```
project/
├── Cargo.toml            # Package manifest
├── Cargo.lock            # Dependency lock (commit for binaries)
├── rust-toolchain.toml   # Rust version pinning
├── .cargo/
│   └── config.toml       # Cargo configuration
├── clippy.toml           # Clippy configuration
├── rustfmt.toml          # Formatter configuration
├── deny.toml             # Dependency checking
└── .github/
    └── workflows/
        └── ci.yml
```

### rust-toolchain.toml
```toml
[toolchain]
channel = "1.75"
components = ["rustfmt", "clippy"]
targets = ["x86_64-unknown-linux-gnu"]
```

### .cargo/config.toml
```toml
[build]
rustflags = ["-D", "warnings"]

[target.x86_64-unknown-linux-gnu]
linker = "clang"
rustflags = ["-C", "link-arg=-fuse-ld=lld"]

[alias]
b = "build"
t = "test"
r = "run"
c = "check"
```

## .gitignore
```gitignore
# Generated
/target/
**/*.rs.bk

# IDE
.idea/
.vscode/
*.swp

# Environment
.env
.env.local

# OS
.DS_Store
Thumbs.db

# Debug
*.pdb
```

## Feature Flags Structure

```toml
# Cargo.toml
[features]
default = ["std"]
std = []
async = ["tokio"]
full = ["std", "async", "serde"]

[dependencies]
tokio = { version = "1", optional = true }
serde = { version = "1", optional = true }
```

```rust
// lib.rs
#[cfg(feature = "async")]
pub mod async_api;

#[cfg(feature = "serde")]
mod serialization;
```

## Conditional Compilation

```rust
// Platform-specific
#[cfg(target_os = "linux")]
mod linux_impl;

#[cfg(target_os = "windows")]
mod windows_impl;

// Feature-gated
#[cfg(feature = "async")]
pub async fn fetch() { }

// Test-only
#[cfg(test)]
mod tests {
    use super::*;
}
```
