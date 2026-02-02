//! XLEX CLI - A streaming Excel manipulation tool.

#![allow(clippy::manual_strip)]
#![allow(clippy::nonminimal_bool)]
#![allow(clippy::double_ended_iterator_last)]
#![allow(clippy::field_reassign_with_default)]
#![allow(clippy::collapsible_else_if)]
#![allow(clippy::collapsible_if)]
#![allow(clippy::unnecessary_lazy_evaluations)]
#![allow(clippy::format_in_format_args)]
#![allow(clippy::print_literal)]
#![allow(clippy::disallowed_names)]

mod commands;
pub mod progress;
// mod config; // TODO: Implement configuration module
// mod output; // TODO: Implement output formatting module

use std::io::Write;
use std::process::ExitCode;

use clap::Parser;
use colored::Colorize;
use commands::Cli;

fn main() -> ExitCode {
    let cli = Cli::parse();

    match cli.run() {
        Ok(()) => ExitCode::SUCCESS,
        Err(e) => {
            // Get error details
            let (exit_code, error_code, suggestion) =
                if let Some(xlex_err) = e.downcast_ref::<xlex_core::XlexError>() {
                    (
                        xlex_err.exit_code(),
                        Some(xlex_err.code().to_string()),
                        xlex_err.recovery_suggestion(),
                    )
                } else {
                    (1, None, None)
                };

            // Log error to file if XLEX_LOG_FILE is set
            if let Ok(log_file) = std::env::var("XLEX_LOG_FILE") {
                log_error_to_file(&log_file, &e, error_code.as_deref());
            }

            // Print error
            if cli.global.json_errors {
                let mut error_json = serde_json::json!({
                    "error": true,
                    "message": e.to_string(),
                    "exit_code": exit_code,
                });
                if let Some(code) = &error_code {
                    error_json["code"] = serde_json::Value::String(code.clone());
                }
                if let Some(hint) = suggestion {
                    error_json["suggestion"] = serde_json::Value::String(hint.to_string());
                }
                eprintln!("{}", serde_json::to_string_pretty(&error_json).unwrap());
            } else {
                if let Some(code) = &error_code {
                    eprintln!("{} [{}]: {}", "error".red().bold(), code.yellow(), e);
                } else {
                    eprintln!("{}: {}", "error".red().bold(), e);
                }

                // Print recovery suggestion if available and not in quiet mode
                if !cli.global.quiet {
                    if let Some(hint) = suggestion {
                        eprintln!("{}: {}", "hint".cyan().bold(), hint);
                    }
                }
            }

            ExitCode::from(exit_code as u8)
        }
    }
}

/// Log error to a file specified by XLEX_LOG_FILE environment variable.
fn log_error_to_file(log_file: &str, error: &anyhow::Error, error_code: Option<&str>) {
    let timestamp = chrono::Local::now().format("%Y-%m-%d %H:%M:%S%.3f");
    let code_str = error_code.unwrap_or("UNKNOWN");

    let log_entry = format!("[{}] [{}] {}\n", timestamp, code_str, error);

    // Append to log file, creating if necessary
    if let Ok(mut file) = std::fs::OpenOptions::new()
        .create(true)
        .append(true)
        .open(log_file)
    {
        let _ = file.write_all(log_entry.as_bytes());
    }
}
