//! Session management for xlex.
//!
//! Provides a persistent session that allows multiple commands to operate
//! on the same workbook without re-opening and re-saving for each command.
//!
//! Workflow:
//!   xlex open report.xlsx        # Creates .xlex/ session
//!   xlex cell set Sheet1 A1 Hi   # Operates on working copy
//!   xlex row append Sheet1 a,b   # Same working copy
//!   xlex status                  # Show session info
//!   xlex commit                  # Save back to original
//!   xlex close                   # Discard changes (alternative to commit)

use std::path::{Path, PathBuf};

use anyhow::Result;
use serde::{Deserialize, Serialize};

/// Session state persisted to `.xlex/session.json`.
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SessionState {
    /// Version of the session format
    pub version: u32,
    /// Absolute path to the original file
    pub original_path: PathBuf,
    /// Absolute path to the working copy
    pub working_path: PathBuf,
    /// When the session was opened
    pub opened_at: String,
}

const SESSION_DIR: &str = ".xlex";
const SESSION_FILE: &str = "session.json";
const WORKING_FILE: &str = "working.xlsx";

/// Returns the session directory path (`.xlex/` in current directory).
pub fn session_dir() -> PathBuf {
    PathBuf::from(SESSION_DIR)
}

/// Returns the session file path.
fn session_file_path() -> PathBuf {
    session_dir().join(SESSION_FILE)
}

/// Check if a session is currently active.
pub fn is_active() -> bool {
    session_file_path().exists()
}

/// Load the current session state, if any.
pub fn load() -> Option<SessionState> {
    let path = session_file_path();
    if !path.exists() {
        return None;
    }
    let content = std::fs::read_to_string(&path).ok()?;
    serde_json::from_str(&content).ok()
}

/// Open a new session for the given file.
pub fn open(file: &Path) -> Result<SessionState> {
    if is_active() {
        let existing = load().unwrap();
        anyhow::bail!(
            "A session is already active for '{}'. Run `xlex close` first or `xlex commit` to save.",
            existing.original_path.display()
        );
    }

    // Resolve to absolute path
    let original_path = std::fs::canonicalize(file)
        .map_err(|e| anyhow::anyhow!("Cannot open '{}': {}", file.display(), e))?;

    // Create .xlex/ directory
    let dir = session_dir();
    std::fs::create_dir_all(&dir)?;

    // Copy file to working copy
    let working_path = std::fs::canonicalize(&dir)?.join(WORKING_FILE);
    std::fs::copy(&original_path, &working_path)?;

    let state = SessionState {
        version: 1,
        original_path,
        working_path,
        opened_at: chrono::Utc::now().to_rfc3339(),
    };

    // Write session state
    let json = serde_json::to_string_pretty(&state)?;
    std::fs::write(session_file_path(), json)?;

    Ok(state)
}

/// Commit the session: copy working copy back to original.
pub fn commit() -> Result<SessionState> {
    let state = load()
        .ok_or_else(|| anyhow::anyhow!("No active session. Run `xlex open <file>` first."))?;

    // Copy working file back to original
    std::fs::copy(&state.working_path, &state.original_path)?;

    // Clean up session
    cleanup()?;

    Ok(state)
}

/// Close the session without saving (discard changes).
pub fn close() -> Result<SessionState> {
    let state = load().ok_or_else(|| anyhow::anyhow!("No active session."))?;

    cleanup()?;

    Ok(state)
}

/// Remove the `.xlex/` directory.
fn cleanup() -> Result<()> {
    let dir = session_dir();
    if dir.exists() {
        std::fs::remove_dir_all(&dir)?;
    }
    Ok(())
}

/// Resolve a file path through the session.
///
/// If a session is active and the given path matches the original file
/// (by name or absolute path), returns the working copy path instead.
/// If no session is active, or the path doesn't match, returns the
/// original path as-is.
#[allow(dead_code)]
pub fn resolve_file(file: &Path) -> PathBuf {
    if let Some(state) = load() {
        // Match by absolute path
        if let Ok(abs) = std::fs::canonicalize(file) {
            if abs == state.original_path {
                return state.working_path;
            }
        }
        // Match by filename (convenience: just type the basename)
        if let Some(file_name) = file.file_name() {
            if let Some(orig_name) = state.original_path.file_name() {
                if file_name == orig_name {
                    return state.working_path;
                }
            }
        }
    }
    file.to_path_buf()
}

#[cfg(test)]
mod tests {
    use super::*;
    use tempfile::TempDir;

    #[test]
    fn test_session_lifecycle() {
        let tmp = TempDir::new().unwrap();
        let original = tmp.path().join("test.xlsx");

        // Create a minimal xlsx for testing
        let wb = xlex_core::Workbook::with_sheets(&["Sheet1"]);
        wb.save_as(&original).unwrap();

        // Change to temp dir so .xlex/ is created there
        let prev_dir = std::env::current_dir().unwrap();
        std::env::set_current_dir(tmp.path()).unwrap();

        // Open session
        let state = open(&original).unwrap();
        assert!(is_active());
        assert!(state.working_path.exists());

        // Double open should fail
        assert!(open(&original).is_err());

        // Close session
        let closed = close().unwrap();
        assert!(!is_active());
        assert_eq!(closed.original_path, state.original_path);

        // Restore dir
        std::env::set_current_dir(prev_dir).unwrap();
    }
}
