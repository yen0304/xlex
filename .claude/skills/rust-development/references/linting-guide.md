# Rust Linting Guide

## Clippy - Official Linter

### Basic Usage
```bash
# Run clippy
cargo clippy

# Check all targets (tests, benches)
cargo clippy --all-targets

# Check all features
cargo clippy --all-features

# Treat warnings as errors (recommended for CI)
cargo clippy -- -D warnings

# Auto-fix issues
cargo clippy --fix
```

### Lint Levels
- `allow` - Ignore
- `warn` - Warning (default)
- `deny` - Error, stops compilation
- `forbid` - Error, cannot be overridden

### Controlling Lints in Code

```rust
// Entire crate
#![allow(clippy::too_many_arguments)]
#![deny(clippy::unwrap_used)]

// Single module or function
#[allow(clippy::needless_return)]
fn example() {
    return 42;
}

// Single line
#[allow(clippy::cast_possible_truncation)]
let x = big_number as u8;
```

### Common Lint Groups

```rust
// Enable strict checks (recommended for new projects)
#![warn(clippy::pedantic)]

// Enable stricter checks
#![warn(clippy::nursery)]

// Restrict unsafe operations
#![warn(clippy::restriction)]

// Performance related
#![warn(clippy::perf)]

// Code complexity
#![warn(clippy::complexity)]

// Code style
#![warn(clippy::style)]
```

### Recommended Lint Configuration

```rust
// Top of lib.rs or main.rs
#![warn(
    clippy::all,
    clippy::pedantic,
    clippy::nursery,
)]
#![allow(
    clippy::module_name_repetitions,  // Allow module::ModuleThing
    clippy::must_use_candidate,       // Don't force #[must_use]
)]

// Production recommended
#![deny(
    clippy::unwrap_used,              // Forbid .unwrap()
    clippy::expect_used,              // Forbid .expect()
    clippy::panic,                    // Forbid panic!
    clippy::todo,                     // Forbid todo!
    clippy::unimplemented,            // Forbid unimplemented!
)]
```

### clippy.toml Configuration

```toml
# clippy.toml (project root)

# Function argument limit
too-many-arguments-threshold = 7

# Cognitive complexity limit
cognitive-complexity-threshold = 25

# Allow unwrap in tests
allow-unwrap-in-tests = true

# Type complexity limit
type-complexity-threshold = 250

# Stack size limit
too-large-for-stack = 200

# Disallowed methods
disallowed-methods = [
    "std::env::var",  # Use custom config system
]

# Disallowed types
disallowed-types = [
    "std::collections::HashMap",  # Use hashbrown
]
```

## Rustfmt - Code Formatter

### Basic Usage
```bash
# Format code
cargo fmt

# Check format (no modify, for CI)
cargo fmt -- --check

# Format single file
rustfmt src/main.rs
```

### rustfmt.toml Configuration

```toml
# rustfmt.toml
edition = "2021"

# Line width
max_width = 100

# Tab settings
tab_spaces = 4
hard_tabs = false

# Import sorting
imports_granularity = "Crate"
group_imports = "StdExternalCrate"
reorder_imports = true

# Function parameters
fn_params_layout = "Tall"

# Use field init shorthand
use_field_init_shorthand = true

# Use try shorthand
use_try_shorthand = true

# Comment formatting
wrap_comments = true
comment_width = 80
normalize_comments = true

# Match arms
match_block_trailing_comma = true

# Structs
struct_lit_single_line = true

# Newline style
newline_style = "Unix"
```

## Rust Analyzer - IDE Integration

### VS Code settings.json
```json
{
    "rust-analyzer.check.command": "clippy",
    "rust-analyzer.check.extraArgs": ["--all-targets"],
    "rust-analyzer.diagnostics.enable": true,
    "rust-analyzer.inlayHints.enable": true,
    "rust-analyzer.lens.enable": true,
    "rust-analyzer.lens.run.enable": true,
    "rust-analyzer.lens.debug.enable": true
}
```

## Other Lint Tools

### cargo-deny - Dependency Checking
```bash
cargo install cargo-deny
cargo deny init  # Generate deny.toml
cargo deny check
```

```toml
# deny.toml
[licenses]
allow = ["MIT", "Apache-2.0", "BSD-3-Clause"]
deny = ["GPL-3.0"]

[bans]
multiple-versions = "warn"
deny = [
    { name = "openssl" },  # Use rustls instead
]

[advisories]
vulnerability = "deny"
unmaintained = "warn"

[sources]
allow-git = []
```

### cargo-audit - Security Vulnerabilities
```bash
cargo install cargo-audit
cargo audit
```

### cargo-outdated - Outdated Dependencies
```bash
cargo install cargo-outdated
cargo outdated
```

### cargo-machete - Unused Dependencies
```bash
cargo install cargo-machete
cargo machete
```

## CI Configuration Examples

### GitHub Actions
```yaml
name: Rust CI

on: [push, pull_request]

jobs:
  lint:
    runs-on: ubuntu-latest
    steps:
      - uses: actions/checkout@v4
      - uses: dtolnay/rust-toolchain@stable
        with:
          components: clippy, rustfmt
      
      - name: Check formatting
        run: cargo fmt -- --check
      
      - name: Clippy
        run: cargo clippy --all-targets --all-features -- -D warnings
      
      - name: Security audit
        run: |
          cargo install cargo-audit
          cargo audit
```

### Pre-commit Hook
```bash
#!/bin/sh
# .git/hooks/pre-commit

set -e

echo "Running cargo fmt..."
cargo fmt -- --check

echo "Running cargo clippy..."
cargo clippy --all-targets -- -D warnings

echo "Running cargo test..."
cargo test --quiet
```

## Common Clippy Warning Fixes

### `clippy::unwrap_used`
```rust
// Bad
let value = map.get("key").unwrap();

// Good
let value = map.get("key").ok_or(MyError::KeyNotFound)?;
// Or
let value = map.get("key").unwrap_or(&default);
```

### `clippy::clone_on_ref_ptr`
```rust
// Bad
let cloned = arc.clone();

// Good
let cloned = Arc::clone(&arc);
```

### `clippy::needless_collect`
```rust
// Bad
let vec: Vec<_> = iter.collect();
for item in vec { ... }

// Good
for item in iter { ... }
```

### `clippy::manual_map`
```rust
// Bad
match opt {
    Some(x) => Some(x + 1),
    None => None,
}

// Good
opt.map(|x| x + 1)
```

### `clippy::redundant_closure`
```rust
// Bad
iter.map(|x| foo(x))

// Good
iter.map(foo)
```
