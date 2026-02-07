use std::collections::HashMap;
use std::path::{Path, PathBuf};
use std::sync::{Arc, Mutex};

use xlex_core::{Workbook, XlexError, XlexResult};

/// Metadata about an open session.
struct SessionEntry {
    workbook: Workbook,
    path: PathBuf,
}

/// Manages open workbook sessions.
#[derive(Clone, Default)]
pub struct SessionStore {
    sessions: Arc<Mutex<HashMap<String, SessionEntry>>>,
}

impl SessionStore {
    pub fn new() -> Self {
        Self::default()
    }

    /// Open a workbook file and create a new session.
    /// Returns the session ID.
    pub fn open(&self, path: &Path) -> XlexResult<String> {
        let workbook = Workbook::open(path)?;
        let session_id = uuid::Uuid::new_v4().to_string();
        let entry = SessionEntry {
            workbook,
            path: path.to_path_buf(),
        };
        let mut sessions = self.sessions.lock().map_err(|e| XlexError::InternalError {
            message: format!("Failed to lock session store: {e}"),
        })?;
        sessions.insert(session_id.clone(), entry);
        Ok(session_id)
    }

    /// Create a new workbook, save it, and open it as a session.
    /// Returns the session ID.
    pub fn create(&self, path: &Path, sheet_names: &[&str]) -> XlexResult<String> {
        let workbook = if sheet_names.is_empty() {
            Workbook::with_sheets(&["Sheet1"])
        } else {
            Workbook::with_sheets(sheet_names)
        };
        workbook.save_as(path)?;
        // Re-open the saved file so the workbook has a proper path
        let workbook = Workbook::open(path)?;
        let session_id = uuid::Uuid::new_v4().to_string();
        let entry = SessionEntry {
            workbook,
            path: path.to_path_buf(),
        };
        let mut sessions = self.sessions.lock().map_err(|e| XlexError::InternalError {
            message: format!("Failed to lock session store: {e}"),
        })?;
        sessions.insert(session_id.clone(), entry);
        Ok(session_id)
    }

    /// Close a session, optionally saving the workbook first.
    pub fn close(&self, session_id: &str, save: bool) -> Result<(), String> {
        let mut sessions = self
            .sessions
            .lock()
            .map_err(|e| format!("Failed to lock session store: {e}"))?;
        let entry = sessions
            .remove(session_id)
            .ok_or_else(|| format!("Session not found: {session_id}"))?;
        if save {
            entry
                .workbook
                .save_as(&entry.path)
                .map_err(|e| format!("Failed to save workbook: {e}"))?;
        }
        Ok(())
    }

    /// Get the file path for a session.
    pub fn get_path(&self, session_id: &str) -> Option<PathBuf> {
        let sessions = self.sessions.lock().ok()?;
        sessions.get(session_id).map(|e| e.path.clone())
    }

    /// Execute a closure with an immutable reference to a workbook.
    pub fn with_workbook<F, R>(&self, session_id: &str, f: F) -> Option<R>
    where
        F: FnOnce(&Workbook) -> R,
    {
        let sessions = self.sessions.lock().ok()?;
        sessions.get(session_id).map(|entry| f(&entry.workbook))
    }

    /// Execute a closure with a mutable reference to a workbook.
    pub fn with_workbook_mut<F, R>(&self, session_id: &str, f: F) -> Option<R>
    where
        F: FnOnce(&mut Workbook, &Path) -> R,
    {
        let mut sessions = self.sessions.lock().ok()?;
        sessions
            .get_mut(session_id)
            .map(|entry| f(&mut entry.workbook, &entry.path))
    }
}
