//! XLEX CLI - A streaming Excel manipulation tool.

mod commands;
// mod config; // TODO: Implement configuration module
// mod output; // TODO: Implement output formatting module

use std::process::ExitCode;

use clap::Parser;
use commands::Cli;

fn main() -> ExitCode {
    let cli = Cli::parse();

    match cli.run() {
        Ok(()) => ExitCode::SUCCESS,
        Err(e) => {
            // Get error code for exit
            let exit_code = if let Some(xlex_err) = e.downcast_ref::<xlex_core::XlexError>() {
                xlex_err.exit_code()
            } else {
                1
            };

            // Print error
            if cli.global.json_errors {
                let error_json = serde_json::json!({
                    "error": true,
                    "message": e.to_string(),
                    "code": exit_code,
                });
                eprintln!("{}", serde_json::to_string_pretty(&error_json).unwrap());
            } else {
                eprintln!("{}: {}", colored::Colorize::red("error"), e);
            }

            ExitCode::from(exit_code as u8)
        }
    }
}
