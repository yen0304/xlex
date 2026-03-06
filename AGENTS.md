# AGENTS.md — Agent Instructions for xlex

## Project Overview

xlex is a high-performance Excel CLI tool written in Rust.

- **Workspace**: 2 crates — `xlex-core` (library) and `xlex-cli` (binary)
- **MSRV**: 1.80
- **Version**: See `Cargo.toml` root

## Development Commands

```bash
cargo test --workspace          # Run all tests (must pass before commit)
cargo clippy --workspace --all-targets -- -D warnings  # Zero warnings required
cargo fmt --all --check         # Formatting check
cargo bench -p xlex-core        # Criterion benchmarks
```

## Pre-commit / Pre-push

Git hooks auto-run tests, fmt check, and clippy. Use `--no-verify` only when CI has already validated.

## CI Notes

- MSRV job generates a fresh lockfile and pins transitive deps (see ci.yml `Generate MSRV-compatible lockfile` step)
- Benchmark job uses **Criterion** (`tool: criterionrs` in github-action-benchmark)
- Docs job runs `mkdocs build --strict` from repo root (`mkdocs.yml` is at root, `docs_dir` defaults to `docs/`)

## Agent Skill Files (IMPORTANT)

Skill files live at `docs/skills/xlex-agent/`:
- `SKILL.md` — Core overview, global flags, quick reference
- `references/commands.md` — Complete CLI command reference
- `references/examples.md` — Real-world workflow examples

**Rule: When adding, removing, or changing any CLI command, subcommand, or flag, you MUST update the corresponding agent skill files to keep them in sync.**

Additionally, the npm package has its own README at `npm/xlex/README.md` with usage examples. When CLI features or commands change, update that README as well to keep it accurate.

## Session Mode (IMPORTANT)

When adding a new read-only CLI feature, you MUST also add it to session mode (`run_session` in `mod.rs`) if the `LazyWorkbook` API supports it. Session mode keeps the file loaded in memory for faster repeated operations — any query command that can work with `stream_rows()` or `read_cell()` should be available there.

**Rule: When implementing a new read/query command, always check if it can be added to session mode and add it if feasible.**

## Release Checklist (IMPORTANT)

When publishing a new version, you MUST update **all** of the following:

1. `Cargo.toml` (workspace root) — bump `version`
2. `npm/xlex/package.json` — match the new version
3. `CHANGELOG.md` — add release entry with summary of changes
4. `docs/skills/xlex-agent/SKILL.md` — update version if shown
5. Git tag — `git tag vX.Y.Z` and push

**Rule: Never push a version tag without first updating CHANGELOG.md, Cargo.toml, and package.json.**
