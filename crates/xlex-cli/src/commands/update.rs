//! Self-update command — checks for new releases and updates the binary.

use anyhow::{Context, Result};
use clap::Parser;

use super::GlobalOptions;

const REPO: &str = "yen0304/xlex";
const CURRENT_VERSION: &str = env!("CARGO_PKG_VERSION");

/// Update xlex to the latest version.
#[derive(Parser, Debug)]
pub struct UpdateArgs {
    /// Check for updates without installing
    #[arg(long)]
    pub check: bool,

    /// Update to a specific version (e.g. v0.3.1)
    #[arg(long)]
    pub target: Option<String>,
}

/// Fetch the latest release tag from GitHub API using curl.
fn fetch_latest_version() -> Result<String> {
    let output = std::process::Command::new("curl")
        .args([
            "-sL",
            &format!("https://api.github.com/repos/{REPO}/releases/latest"),
        ])
        .output()
        .context("Failed to run curl — is it installed?")?;

    if !output.status.success() {
        anyhow::bail!(
            "GitHub API request failed: {}",
            String::from_utf8_lossy(&output.stderr)
        );
    }

    let body = String::from_utf8_lossy(&output.stdout);

    // Parse tag_name from JSON (lightweight — no serde needed for one field)
    // Expected format: `  "tag_name": "v0.3.1",`
    let tag = body
        .lines()
        .find(|l| l.contains("\"tag_name\""))
        .and_then(|l| {
            let colon = l.find(':')?;
            let after_colon = &l[colon + 1..];
            let first_quote = after_colon.find('"')? + 1;
            let value_start = &after_colon[first_quote..];
            let end_quote = value_start.find('"')?;
            Some(value_start[..end_quote].to_string())
        })
        .context("Could not parse latest version from GitHub API response")?;

    Ok(tag)
}

/// Normalize version string: strip leading 'v' for comparison.
fn normalize(v: &str) -> &str {
    v.strip_prefix('v').unwrap_or(v)
}

pub fn run(args: &UpdateArgs, global: &GlobalOptions) -> Result<()> {
    let current = normalize(CURRENT_VERSION);

    if !global.quiet {
        println!("Current version: v{current}");
        println!("Checking for updates...");
    }

    let target_tag = if let Some(ref v) = args.target {
        if !v.starts_with('v') {
            format!("v{v}")
        } else {
            v.clone()
        }
    } else {
        fetch_latest_version()?
    };
    let latest = normalize(&target_tag);

    if current == latest {
        if !global.quiet {
            println!("Already up to date (v{current}).");
        }
        return Ok(());
    }

    if !global.quiet {
        println!("New version available: v{latest}");
    }

    if args.check {
        // --check: just report, don't install
        if global.format == super::OutputFormat::Json {
            println!(r#"{{"current":"v{current}","latest":"v{latest}","update_available":true}}"#);
        }
        return Ok(());
    }

    if global.dry_run {
        println!("Would update from v{current} to v{latest}");
        return Ok(());
    }

    // Run the install script with the target version
    if !global.quiet {
        println!("Updating to v{latest}...");
    }

    let install_url = format!("https://raw.githubusercontent.com/{REPO}/main/install.sh");

    let status = std::process::Command::new("bash")
        .args([
            "-c",
            &format!("curl -fsSL '{install_url}' | XLEX_VERSION='{target_tag}' bash"),
        ])
        .status()
        .context("Failed to run install script")?;

    if !status.success() {
        anyhow::bail!("Update failed (install script exited with {})", status);
    }

    if !global.quiet {
        println!("Successfully updated to v{latest}!");
    }

    Ok(())
}
