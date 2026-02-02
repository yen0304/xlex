---
name: rust-development
description: Professional Rust development workflows including project setup, error handling, async programming, testing, performance optimization, and common patterns. Use when working with Rust code, Cargo projects, implementing traits, handling Results/Options, writing async code, debugging borrow checker issues, or optimizing Rust performance.
---

# Rust Development

## Project Setup

### New Project
```bash
cargo new project-name        # Binary
cargo new --lib library-name  # Library
```

### Cargo.toml Essentials
```toml
[package]
name = "project"
version = "0.1.0"
edition = "2021"
rust-version = "1.75"

[dependencies]
tokio = { version = "1", features = ["full"] }
serde = { version = "1", features = ["derive"] }
anyhow = "1"
thiserror = "1"

[dev-dependencies]
criterion = "0.5"

[[bench]]
name = "benchmarks"
harness = false

[profile.release]
lto = true
codegen-units = 1
```

## Error Handling

### Custom Errors with thiserror
```rust
use thiserror::Error;

#[derive(Error, Debug)]
pub enum AppError {
    #[error("Database error: {0}")]
    Database(#[from] sqlx::Error),
    #[error("Invalid input: {field}")]
    Validation { field: String },
    #[error("Not found: {0}")]
    NotFound(String),
}

pub type Result<T> = std::result::Result<T, AppError>;
```

### Result/Option Chaining
```rust
// Prefer ? operator
fn process() -> Result<Data> {
    let config = load_config()?;
    let data = fetch_data(&config)?;
    Ok(transform(data))
}

// Option combinators
let value = map.get("key")
    .filter(|v| !v.is_empty())
    .map(|v| v.to_uppercase())
    .unwrap_or_default();
```

## Async Programming

### Tokio Runtime Setup
```rust
#[tokio::main]
async fn main() -> anyhow::Result<()> {
    // Application code
    Ok(())
}

// Custom runtime
fn main() {
    let rt = tokio::runtime::Builder::new_multi_thread()
        .worker_threads(4)
        .enable_all()
        .build()
        .unwrap();
    rt.block_on(async_main());
}
```

### Concurrent Tasks
```rust
use tokio::task::JoinSet;

async fn parallel_fetch(urls: Vec<String>) -> Vec<Response> {
    let mut set = JoinSet::new();
    for url in urls {
        set.spawn(async move { fetch(&url).await });
    }
    
    let mut results = Vec::new();
    while let Some(res) = set.join_next().await {
        if let Ok(data) = res {
            results.push(data);
        }
    }
    results
}
```

### Channels
```rust
use tokio::sync::mpsc;

let (tx, mut rx) = mpsc::channel(100);

tokio::spawn(async move {
    while let Some(msg) = rx.recv().await {
        process(msg).await;
    }
});

tx.send(message).await?;
```

## Common Patterns

### Builder Pattern
```rust
#[derive(Default)]
pub struct Config {
    host: String,
    port: u16,
    timeout: Duration,
}

impl Config {
    pub fn builder() -> ConfigBuilder {
        ConfigBuilder::default()
    }
}

#[derive(Default)]
pub struct ConfigBuilder {
    host: Option<String>,
    port: Option<u16>,
    timeout: Option<Duration>,
}

impl ConfigBuilder {
    pub fn host(mut self, host: impl Into<String>) -> Self {
        self.host = Some(host.into());
        self
    }
    
    pub fn port(mut self, port: u16) -> Self {
        self.port = Some(port);
        self
    }
    
    pub fn build(self) -> Result<Config, &'static str> {
        Ok(Config {
            host: self.host.ok_or("host required")?,
            port: self.port.unwrap_or(8080),
            timeout: self.timeout.unwrap_or(Duration::from_secs(30)),
        })
    }
}
```

### Newtype Pattern
```rust
pub struct UserId(pub i64);
pub struct Email(String);

impl Email {
    pub fn new(s: &str) -> Result<Self, &'static str> {
        if s.contains('@') {
            Ok(Self(s.to_string()))
        } else {
            Err("invalid email")
        }
    }
}
```

### Type State Pattern
```rust
struct Request<S> {
    state: S,
    data: Vec<u8>,
}

struct Building;
struct Ready;

impl Request<Building> {
    fn new() -> Self {
        Self { state: Building, data: vec![] }
    }
    
    fn add_data(mut self, data: &[u8]) -> Self {
        self.data.extend(data);
        self
    }
    
    fn finalize(self) -> Request<Ready> {
        Request { state: Ready, data: self.data }
    }
}

impl Request<Ready> {
    fn send(self) -> Response { /* ... */ }
}
```

## Testing

### Unit Tests
```rust
#[cfg(test)]
mod tests {
    use super::*;
    
    #[test]
    fn test_basic() {
        assert_eq!(add(2, 2), 4);
    }
    
    #[test]
    #[should_panic(expected = "division by zero")]
    fn test_panic() {
        divide(1, 0);
    }
}
```

### Async Tests
```rust
#[tokio::test]
async fn test_async_operation() {
    let result = fetch_data().await;
    assert!(result.is_ok());
}
```

### Integration Tests
Place in `tests/` directory:
```rust
// tests/integration_test.rs
use my_crate::public_api;

#[test]
fn test_public_interface() {
    let result = public_api::process("input");
    assert_eq!(result, expected);
}
```

### Property-Based Testing
```rust
use proptest::prelude::*;

proptest! {
    #[test]
    fn test_roundtrip(s in "\\PC*") {
        let encoded = encode(&s);
        let decoded = decode(&encoded)?;
        prop_assert_eq!(s, decoded);
    }
}
```

## Borrow Checker Solutions

### Common Fixes

**Multiple mutable borrows:**
```rust
// Problem: Cannot borrow twice
// Solution: Use indices or split_at_mut
let (left, right) = slice.split_at_mut(mid);
```

**Self-referential structs:**
```rust
// Use ouroboros or self_cell crate
use ouroboros::self_referencing;

#[self_referencing]
struct SelfRef {
    data: String,
    #[borrows(data)]
    slice: &'this str,
}
```

**Interior mutability:**
```rust
use std::cell::RefCell;
use std::rc::Rc;

// Single-threaded
let shared = Rc::new(RefCell::new(data));

// Multi-threaded
use std::sync::{Arc, Mutex};
let shared = Arc::new(Mutex::new(data));
```

## Performance

### Profiling
```bash
# CPU profiling
cargo install flamegraph
cargo flamegraph --bin my-app

# Memory profiling
cargo install cargo-instruments  # macOS
cargo instruments -t Allocations
```

### Optimization Tips
```rust
// Preallocate collections
let mut vec = Vec::with_capacity(expected_size);

// Avoid cloning in loops
for item in &items {  // Borrow, don't move
    process(item);
}

// Use iterators over indexing
items.iter()
    .filter(|x| x.is_valid())
    .map(|x| x.transform())
    .collect()

// Inline hot functions
#[inline]
fn hot_path() { /* ... */ }
```

### Benchmarking with Criterion
```rust
// benches/benchmarks.rs
use criterion::{criterion_group, criterion_main, Criterion};

fn benchmark_function(c: &mut Criterion) {
    c.bench_function("my_function", |b| {
        b.iter(|| my_function(test_input))
    });
}

criterion_group!(benches, benchmark_function);
criterion_main!(benches);
```

Run: `cargo bench`

## CLI Development

### Clap for Arguments
```rust
use clap::Parser;

#[derive(Parser)]
#[command(name = "app", version, about)]
struct Cli {
    /// Input file path
    #[arg(short, long)]
    input: PathBuf,
    
    /// Verbose output
    #[arg(short, long, action = clap::ArgAction::Count)]
    verbose: u8,
    
    #[command(subcommand)]
    command: Commands,
}

#[derive(clap::Subcommand)]
enum Commands {
    /// Process files
    Process { files: Vec<PathBuf> },
    /// Show config
    Config,
}
```

## Web Development

### Axum Server
```rust
use axum::{routing::get, Router, Json};
use serde::Serialize;

#[derive(Serialize)]
struct Health { status: String }

async fn health() -> Json<Health> {
    Json(Health { status: "ok".into() })
}

#[tokio::main]
async fn main() {
    let app = Router::new()
        .route("/health", get(health));
    
    let listener = tokio::net::TcpListener::bind("0.0.0.0:3000")
        .await.unwrap();
    axum::serve(listener, app).await.unwrap();
}
```

## Useful Commands

```bash
cargo check              # Fast syntax check
cargo clippy             # Linting
cargo fmt                # Format code
cargo doc --open         # Generate docs
cargo tree               # Dependency tree
cargo update             # Update deps
cargo audit              # Security audit
cargo expand             # Expand macros
RUST_BACKTRACE=1 cargo run  # Debug panics
```

## Linting

### Quick Clippy Setup
```rust
// 在 lib.rs 或 main.rs 頂部
#![warn(clippy::all, clippy::pedantic, clippy::nursery)]
#![deny(clippy::unwrap_used, clippy::expect_used)]
```

### CI 必備命令
```bash
cargo fmt -- --check           # 格式檢查
cargo clippy -- -D warnings    # Lint 檢查
cargo audit                    # 安全漏洞
```

For detailed linting configuration, see [references/linting-guide.md](references/linting-guide.md)

## Additional Resources

- For project structure and workspace setup, see [references/project-structure.md](references/project-structure.md)
- For serde patterns and custom serialization, see [references/serde-patterns.md](references/serde-patterns.md)
- For unsafe Rust guidelines, see [references/unsafe-guide.md](references/unsafe-guide.md)
- For FFI and C interop, see [references/ffi-guide.md](references/ffi-guide.md)
- For Clippy, rustfmt, and CI linting setup, see [references/linting-guide.md](references/linting-guide.md)
