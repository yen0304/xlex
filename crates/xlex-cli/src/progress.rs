//! Progress indicators for long-running operations.
//!
//! This module provides progress bars and spinners for CLI operations
//! that may take a long time, such as importing/exporting large files.

use indicatif::{ProgressBar, ProgressStyle};
use std::time::Duration;

/// Progress indicator types.
pub enum ProgressKind {
    /// Spinner for operations with unknown total
    Spinner,
    /// Progress bar with known total
    Bar { total: u64 },
}

/// A wrapper around indicatif progress bar for consistent styling.
pub struct Progress {
    bar: ProgressBar,
    quiet: bool,
}

impl Progress {
    /// Create a new progress indicator.
    ///
    /// If `quiet` is true, the progress bar will be hidden.
    pub fn new(kind: ProgressKind, message: &str, quiet: bool) -> Self {
        let bar = match kind {
            ProgressKind::Spinner => {
                let pb = ProgressBar::new_spinner();
                pb.set_style(
                    ProgressStyle::default_spinner()
                        .tick_chars("⠋⠙⠹⠸⠼⠴⠦⠧⠇⠏")
                        .template("{spinner:.cyan} {msg}")
                        .unwrap(),
                );
                pb.set_message(message.to_string());
                pb.enable_steady_tick(Duration::from_millis(100));
                pb
            }
            ProgressKind::Bar { total } => {
                let pb = ProgressBar::new(total);
                pb.set_style(
                    ProgressStyle::default_bar()
                        .template("{spinner:.cyan} {msg} [{bar:40.cyan/blue}] {pos}/{len} ({eta})")
                        .unwrap()
                        .progress_chars("█▓▒░"),
                );
                pb.set_message(message.to_string());
                pb
            }
        };

        if quiet {
            bar.set_draw_target(indicatif::ProgressDrawTarget::hidden());
        }

        Self { bar, quiet }
    }

    /// Create a spinner for an operation with unknown total.
    pub fn spinner(message: &str, quiet: bool) -> Self {
        Self::new(ProgressKind::Spinner, message, quiet)
    }

    /// Create a progress bar with a known total.
    pub fn bar(total: u64, message: &str, quiet: bool) -> Self {
        Self::new(ProgressKind::Bar { total }, message, quiet)
    }

    /// Update the progress message.
    pub fn set_message(&self, message: &str) {
        self.bar.set_message(message.to_string());
    }

    /// Increment the progress by a given amount.
    pub fn inc(&self, delta: u64) {
        self.bar.inc(delta);
    }

    /// Set the current position.
    pub fn set_position(&self, pos: u64) {
        self.bar.set_position(pos);
    }

    /// Set the total length (for progress bars).
    pub fn set_length(&self, len: u64) {
        self.bar.set_length(len);
    }

    /// Tick the progress (for spinners).
    pub fn tick(&self) {
        self.bar.tick();
    }

    /// Mark the progress as finished with a success message.
    pub fn finish_with_message(&self, message: &str) {
        self.bar.finish_with_message(message.to_string());
    }

    /// Mark the progress as finished and clear.
    pub fn finish_and_clear(&self) {
        self.bar.finish_and_clear();
    }

    /// Check if the progress bar is hidden (quiet mode).
    pub fn is_quiet(&self) -> bool {
        self.quiet
    }

    /// Abandon the progress bar (on error).
    pub fn abandon_with_message(&self, message: &str) {
        self.bar.abandon_with_message(message.to_string());
    }
}

/// A multi-progress manager for operations with multiple progress bars.
pub struct MultiProgress {
    multi: indicatif::MultiProgress,
    quiet: bool,
}

impl MultiProgress {
    /// Create a new multi-progress manager.
    pub fn new(quiet: bool) -> Self {
        let multi = indicatif::MultiProgress::new();
        if quiet {
            multi.set_draw_target(indicatif::ProgressDrawTarget::hidden());
        }
        Self { multi, quiet }
    }

    /// Add a spinner to the multi-progress.
    pub fn add_spinner(&self, message: &str) -> Progress {
        let pb = ProgressBar::new_spinner();
        pb.set_style(
            ProgressStyle::default_spinner()
                .tick_chars("⠋⠙⠹⠸⠼⠴⠦⠧⠇⠏")
                .template("{spinner:.cyan} {msg}")
                .unwrap(),
        );
        pb.set_message(message.to_string());
        pb.enable_steady_tick(Duration::from_millis(100));
        let bar = self.multi.add(pb);
        Progress {
            bar,
            quiet: self.quiet,
        }
    }

    /// Add a progress bar to the multi-progress.
    pub fn add_bar(&self, total: u64, message: &str) -> Progress {
        let pb = ProgressBar::new(total);
        pb.set_style(
            ProgressStyle::default_bar()
                .template("{spinner:.cyan} {msg} [{bar:40.cyan/blue}] {pos}/{len} ({eta})")
                .unwrap()
                .progress_chars("█▓▒░"),
        );
        pb.set_message(message.to_string());
        let bar = self.multi.add(pb);
        Progress {
            bar,
            quiet: self.quiet,
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_spinner_creation() {
        let progress = Progress::spinner("Loading...", true);
        assert!(progress.is_quiet());
    }

    #[test]
    fn test_bar_creation() {
        let progress = Progress::bar(100, "Processing...", true);
        progress.inc(50);
        progress.finish_with_message("Done");
    }

    #[test]
    fn test_multi_progress() {
        let multi = MultiProgress::new(true);
        let spinner = multi.add_spinner("Loading...");
        let bar = multi.add_bar(100, "Processing...");

        spinner.tick();
        bar.inc(10);
    }

    #[test]
    fn test_progress_set_message() {
        let progress = Progress::spinner("Initial", true);
        progress.set_message("Updated");
        assert!(progress.is_quiet());
    }

    #[test]
    fn test_progress_set_position() {
        let progress = Progress::bar(100, "Test", true);
        progress.set_position(50);
        progress.inc(10);
    }

    #[test]
    fn test_progress_set_length() {
        let progress = Progress::bar(50, "Test", true);
        progress.set_length(100);
        progress.inc(25);
    }

    #[test]
    fn test_progress_tick() {
        let progress = Progress::spinner("Tick test", true);
        progress.tick();
        progress.tick();
        progress.finish_with_message("Done");
    }

    #[test]
    fn test_progress_finish_and_clear() {
        let progress = Progress::spinner("Clear test", true);
        progress.tick();
        progress.finish_and_clear();
    }

    #[test]
    fn test_progress_abandon_with_message() {
        let progress = Progress::bar(100, "Abandon test", true);
        progress.inc(30);
        progress.abandon_with_message("Failed");
    }

    #[test]
    fn test_progress_kind_spinner() {
        let progress = Progress::new(ProgressKind::Spinner, "Spinner", true);
        progress.tick();
        assert!(progress.is_quiet());
    }

    #[test]
    fn test_progress_kind_bar() {
        let progress = Progress::new(ProgressKind::Bar { total: 100 }, "Bar", true);
        progress.inc(50);
        assert!(progress.is_quiet());
    }

    #[test]
    fn test_multi_progress_quiet() {
        let multi = MultiProgress::new(true);
        let spinner = multi.add_spinner("Test");
        assert!(spinner.is_quiet());
    }

    #[test]
    fn test_multi_progress_bar_operations() {
        let multi = MultiProgress::new(true);
        let bar = multi.add_bar(50, "Test bar");
        bar.set_position(25);
        bar.set_length(100);
        bar.inc(10);
        bar.finish_with_message("Complete");
    }

    #[test]
    fn test_spinner_non_quiet() {
        // Test non-quiet mode (still creates progress bar but visible)
        let progress = Progress::spinner("Visible", false);
        assert!(!progress.is_quiet());
        progress.finish_and_clear();
    }

    #[test]
    fn test_bar_non_quiet() {
        let progress = Progress::bar(100, "Visible bar", false);
        assert!(!progress.is_quiet());
        progress.finish_and_clear();
    }

    #[test]
    fn test_multi_progress_non_quiet() {
        let multi = MultiProgress::new(false);
        let spinner = multi.add_spinner("Visible spinner");
        let bar = multi.add_bar(100, "Visible bar");
        assert!(!spinner.is_quiet());
        assert!(!bar.is_quiet());
        spinner.finish_and_clear();
        bar.finish_and_clear();
    }
}
