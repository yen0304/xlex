//! CLI commands for xlex.

mod cell;
mod column;
mod export;
mod formula;
mod import;
mod range;
mod row;
mod sheet;
mod style;
mod template;
mod workbook;

use anyhow::Result;
use clap::{Parser, Subcommand};

/// XLEX - A streaming Excel manipulation tool.
///
/// XLEX provides CLI-first, streaming-based Excel manipulation for developers
/// and automation pipelines. It can handle files up to 200MB without memory
/// exhaustion.
#[derive(Parser)]
#[command(name = "xlex")]
#[command(author, version, about, long_about = None)]
#[command(propagate_version = true)]
pub struct Cli {
    #[command(subcommand)]
    pub command: Commands,

    #[command(flatten)]
    pub global: GlobalOptions,
}

/// Global options available for all commands.
#[derive(Parser, Debug, Clone)]
pub struct GlobalOptions {
    /// Suppress all output except errors
    #[arg(long, short = 'q', global = true)]
    pub quiet: bool,

    /// Enable verbose output
    #[arg(long, short = 'v', global = true)]
    pub verbose: bool,

    /// Output format (text, json, csv)
    #[arg(long, short = 'f', global = true, default_value = "text")]
    pub format: OutputFormat,

    /// Disable colored output
    #[arg(long, global = true)]
    pub no_color: bool,

    /// Force colored output even when piped
    #[arg(long, global = true)]
    pub color: bool,

    /// Output errors as JSON
    #[arg(long, global = true)]
    pub json_errors: bool,

    /// Perform a dry run without making changes
    #[arg(long, global = true)]
    pub dry_run: bool,

    /// Write output to file instead of stdout
    #[arg(long, short = 'o', global = true)]
    pub output: Option<std::path::PathBuf>,
}

/// Output format options.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq, clap::ValueEnum)]
pub enum OutputFormat {
    #[default]
    Text,
    Json,
    Csv,
    Ndjson,
}

/// Available subcommands.
#[derive(Subcommand)]
pub enum Commands {
    // Workbook operations
    /// Display workbook information
    Info(workbook::InfoArgs),
    /// Validate workbook structure
    Validate(workbook::ValidateArgs),
    /// Create a copy of a workbook
    Clone(workbook::CloneArgs),
    /// Create a new workbook
    Create(workbook::CreateArgs),
    /// Get or set workbook properties
    Props(workbook::PropsArgs),
    /// Display workbook statistics
    Stats(workbook::StatsArgs),

    // Sheet operations
    /// Sheet operations (list, add, remove, rename, etc.)
    Sheet(sheet::SheetArgs),

    // Cell operations
    /// Cell operations (get, set, clear, etc.)
    Cell(cell::CellArgs),

    // Row operations
    /// Row operations (get, append, insert, delete, etc.)
    Row(row::RowArgs),

    // Column operations
    /// Column operations (get, insert, delete, etc.)
    Column(column::ColumnArgs),

    // Range operations
    /// Range operations (get, copy, move, merge, etc.)
    Range(range::RangeArgs),

    // Style operations
    /// Style operations (list, get, apply, etc.)
    Style(style::StyleArgs),

    // Formula operations
    /// Formula operations (get, set, list, check, etc.)
    Formula(formula::FormulaArgs),

    // Template operations
    /// Template operations (apply, list, validate, etc.)
    Template(template::TemplateArgs),

    // Import/Export
    /// Import from external format
    Import(import::ImportArgs),

    /// Export to external format
    Export(export::ExportArgs),

    /// Convert between formats
    Convert(ConvertArgs),

    // Utility commands
    /// Generate shell completion scripts
    Completion(CompletionArgs),

    /// Display or modify configuration
    Config(ConfigArgs),

    /// Execute batch commands from file or stdin
    Batch(BatchArgs),

    /// Manage command aliases
    Alias(AliasArgs),

    /// Display version information
    Version,
}

/// Convert arguments.
#[derive(Parser)]
pub struct ConvertArgs {
    /// Input file
    pub input: std::path::PathBuf,
    /// Output file
    pub output: std::path::PathBuf,
}

/// Shell completion arguments.
#[derive(Parser)]
pub struct CompletionArgs {
    /// Shell to generate completions for
    #[arg(value_enum)]
    pub shell: clap_complete::Shell,
}

/// Batch execution arguments.
#[derive(Parser)]
pub struct BatchArgs {
    /// Read commands from file (use - for stdin)
    #[arg(short, long)]
    pub file: Option<std::path::PathBuf>,
    /// Continue executing on error
    #[arg(long)]
    pub continue_on_error: bool,
}

/// Alias management arguments.
#[derive(Parser)]
pub struct AliasArgs {
    #[command(subcommand)]
    pub command: AliasCommand,
}

#[derive(Subcommand)]
pub enum AliasCommand {
    /// List all aliases
    List,
    /// Add a new alias
    Add {
        /// Alias name
        name: String,
        /// Command to alias
        command: String,
    },
    /// Remove an alias
    Remove {
        /// Alias name
        name: String,
    },
}

/// Configuration arguments.
#[derive(Parser)]
pub struct ConfigArgs {
    #[command(subcommand)]
    pub command: ConfigCommand,
}

#[derive(Subcommand)]
pub enum ConfigCommand {
    /// Show current configuration
    Show {
        /// Show effective config from all sources
        #[arg(long)]
        effective: bool,
    },
    /// Get a configuration value
    Get {
        /// Configuration key
        key: String,
    },
    /// Set a configuration value
    Set {
        /// Configuration key
        key: String,
        /// Configuration value
        value: String,
    },
    /// Reset configuration to defaults
    Reset,
    /// Initialize configuration file
    Init,
    /// Validate configuration file
    Validate,
}

impl Cli {
    /// Runs the CLI command.
    pub fn run(&self) -> Result<()> {
        // Set up colored output
        if self.global.no_color {
            colored::control::set_override(false);
        } else if self.global.color {
            colored::control::set_override(true);
        }

        match &self.command {
            // Workbook operations
            Commands::Info(args) => workbook::info(args, &self.global),
            Commands::Validate(args) => workbook::validate(args, &self.global),
            Commands::Clone(args) => workbook::clone(args, &self.global),
            Commands::Create(args) => workbook::create(args, &self.global),
            Commands::Props(args) => workbook::props(args, &self.global),
            Commands::Stats(args) => workbook::stats(args, &self.global),

            // Sheet operations
            Commands::Sheet(args) => sheet::run(args, &self.global),

            // Cell operations
            Commands::Cell(args) => cell::run(args, &self.global),

            // Row operations
            Commands::Row(args) => row::run(args, &self.global),

            // Column operations
            Commands::Column(args) => column::run(args, &self.global),

            // Range operations
            Commands::Range(args) => range::run(args, &self.global),

            // Style operations
            Commands::Style(args) => style::run(args, &self.global),

            // Formula operations
            Commands::Formula(args) => formula::run(args, &self.global),

            // Template operations
            Commands::Template(args) => template::run(args, &self.global),

            // Import/Export
            Commands::Export(args) => export::run(args, &self.global),
            Commands::Import(args) => import::run(args, &self.global),
            Commands::Convert(args) => run_convert(args, &self.global),

            // Utility
            Commands::Completion(args) => run_completion(args),
            Commands::Config(args) => run_config(args, &self.global),
            Commands::Batch(args) => run_batch(args, &self.global),
            Commands::Alias(args) => run_alias(args, &self.global),
            Commands::Version => run_version(&self.global),
        }
    }
}

fn run_convert(args: &ConvertArgs, global: &GlobalOptions) -> Result<()> {
    use xlex_core::Workbook;

    let input = &args.input;
    let output = &args.output;

    // Determine formats from extensions
    let input_ext = input
        .extension()
        .map(|e| e.to_string_lossy().to_lowercase())
        .unwrap_or_default();
    let output_ext = output
        .extension()
        .map(|e| e.to_string_lossy().to_lowercase())
        .unwrap_or_default();

    if global.dry_run {
        println!(
            "Would convert {} ({}) to {} ({})",
            input.display(),
            input_ext,
            output.display(),
            output_ext
        );
        return Ok(());
    }

    // Handle conversion based on input and output formats
    match (input_ext.as_str(), output_ext.as_str()) {
        // CSV -> XLSX
        ("csv", "xlsx") => {
            let import_args = import::ImportArgs {
                command: import::ImportCommand::Csv {
                    source: input.clone(),
                    dest: output.clone(),
                    sheet: None,
                    delimiter: ',',
                    header: false,
                },
            };
            import::run(&import_args, global)
        }
        // TSV -> XLSX
        ("tsv", "xlsx") => {
            let import_args = import::ImportArgs {
                command: import::ImportCommand::Tsv {
                    source: input.clone(),
                    dest: output.clone(),
                    sheet: None,
                },
            };
            import::run(&import_args, global)
        }
        // JSON -> XLSX
        ("json", "xlsx") => {
            let import_args = import::ImportArgs {
                command: import::ImportCommand::Json {
                    source: input.clone(),
                    dest: output.clone(),
                    sheet: None,
                },
            };
            import::run(&import_args, global)
        }
        // NDJSON -> XLSX
        ("ndjson", "xlsx") => {
            let import_args = import::ImportArgs {
                command: import::ImportCommand::Ndjson {
                    source: input.clone(),
                    dest: output.clone(),
                    sheet: None,
                    header: true,
                },
            };
            import::run(&import_args, global)
        }
        // XLSX -> CSV
        ("xlsx", "csv") => {
            let export_args = export::ExportArgs {
                command: export::ExportCommand::Csv {
                    source: input.clone(),
                    dest: output.to_string_lossy().to_string(),
                    sheet: None,
                    delimiter: ',',
                    all: false,
                },
            };
            export::run(&export_args, global)
        }
        // XLSX -> TSV
        ("xlsx", "tsv") => {
            let export_args = export::ExportArgs {
                command: export::ExportCommand::Tsv {
                    source: input.clone(),
                    dest: output.to_string_lossy().to_string(),
                    sheet: None,
                    all: false,
                },
            };
            export::run(&export_args, global)
        }
        // XLSX -> JSON
        ("xlsx", "json") => {
            let export_args = export::ExportArgs {
                command: export::ExportCommand::Json {
                    source: input.clone(),
                    dest: output.to_string_lossy().to_string(),
                    sheet: None,
                    header: true,
                    all: false,
                },
            };
            export::run(&export_args, global)
        }
        // XLSX -> NDJSON
        ("xlsx", "ndjson") => {
            let export_args = export::ExportArgs {
                command: export::ExportCommand::Ndjson {
                    source: input.clone(),
                    dest: output.to_string_lossy().to_string(),
                    sheet: None,
                    header: true,
                    all: false,
                },
            };
            export::run(&export_args, global)
        }
        // XLSX -> YAML
        ("xlsx", "yaml") | ("xlsx", "yml") => {
            let export_args = export::ExportArgs {
                command: export::ExportCommand::Yaml {
                    source: input.clone(),
                    dest: output.to_string_lossy().to_string(),
                    sheet: None,
                    all: false,
                },
            };
            export::run(&export_args, global)
        }
        // XLSX -> MD
        ("xlsx", "md") => {
            let export_args = export::ExportArgs {
                command: export::ExportCommand::Markdown {
                    source: input.clone(),
                    dest: output.to_string_lossy().to_string(),
                    sheet: None,
                    all: false,
                },
            };
            export::run(&export_args, global)
        }
        // XLSX -> XLSX (copy)
        ("xlsx", "xlsx") => {
            let workbook = Workbook::open(input)?;
            workbook.save_as(output)?;
            if !global.quiet {
                println!("Copied {} to {}", input.display(), output.display());
            }
            Ok(())
        }
        _ => anyhow::bail!(
            "Unsupported conversion: {} -> {}. Supported: csv, tsv, json, ndjson, yaml, md <-> xlsx",
            input_ext,
            output_ext
        ),
    }
}

fn run_completion(args: &CompletionArgs) -> Result<()> {
    let mut cmd = Cli::command();
    clap_complete::generate(args.shell, &mut cmd, "xlex", &mut std::io::stdout());
    Ok(())
}

fn get_config_dir() -> Result<std::path::PathBuf> {
    dirs::config_dir()
        .map(|p| p.join("xlex"))
        .ok_or_else(|| anyhow::anyhow!("Could not determine config directory"))
}

fn get_config_path() -> Result<std::path::PathBuf> {
    Ok(get_config_dir()?.join("config.yml"))
}

fn run_config(args: &ConfigArgs, global: &GlobalOptions) -> Result<()> {
    use colored::Colorize;

    match &args.command {
        ConfigCommand::Show { effective } => {
            let config_path = get_config_path()?;
            
            if global.format == OutputFormat::Json {
                let config = if config_path.exists() {
                    std::fs::read_to_string(&config_path)?
                } else {
                    "{}".to_string()
                };
                println!("{}", config);
            } else {
                println!("{}: {}\n", "Config file".bold(), config_path.display());
                
                if config_path.exists() {
                    let content = std::fs::read_to_string(&config_path)?;
                    println!("{}", content);
                } else {
                    println!("{}", "(No config file found. Run 'xlex config init' to create one.)".dimmed());
                }

                if *effective {
                    println!("\n{}:", "Effective values".bold());
                    println!("  default_format: {}", global.format as u8);
                    println!("  quiet: {}", global.quiet);
                    println!("  verbose: {}", global.verbose);
                    println!("  no_color: {}", global.no_color);
                }
            }
        }
        ConfigCommand::Get { key } => {
            let config_path = get_config_path()?;
            if !config_path.exists() {
                anyhow::bail!("No config file found");
            }

            let content = std::fs::read_to_string(&config_path)?;
            let yaml: serde_yaml::Value = serde_yaml::from_str(&content)?;

            if let Some(value) = yaml.get(key) {
                println!("{}", serde_yaml::to_string(value)?.trim());
            } else {
                anyhow::bail!("Key '{}' not found in config", key);
            }
        }
        ConfigCommand::Set { key, value } => {
            if global.dry_run {
                println!("Would set {} = {}", key, value);
                return Ok(());
            }

            let config_path = get_config_path()?;
            let config_dir = get_config_dir()?;

            // Ensure config directory exists
            std::fs::create_dir_all(&config_dir)?;

            // Load existing or create new config
            let mut yaml: serde_yaml::Mapping = if config_path.exists() {
                let content = std::fs::read_to_string(&config_path)?;
                serde_yaml::from_str(&content).unwrap_or_default()
            } else {
                serde_yaml::Mapping::new()
            };

            // Set the value
            yaml.insert(
                serde_yaml::Value::String(key.clone()),
                serde_yaml::Value::String(value.clone()),
            );

            // Write back
            std::fs::write(&config_path, serde_yaml::to_string(&yaml)?)?;

            if !global.quiet {
                println!("{} Set {} = {}", "✓".green(), key.cyan(), value);
            }
        }
        ConfigCommand::Reset => {
            if global.dry_run {
                println!("Would reset config to defaults");
                return Ok(());
            }

            let config_path = get_config_path()?;
            if config_path.exists() {
                std::fs::remove_file(&config_path)?;
                if !global.quiet {
                    println!("{} Config reset to defaults", "✓".green());
                }
            } else {
                println!("No config file to reset");
            }
        }
        ConfigCommand::Init => {
            if global.dry_run {
                println!("Would create config file");
                return Ok(());
            }

            let config_path = get_config_path()?;
            if config_path.exists() {
                anyhow::bail!("Config file already exists: {}", config_path.display());
            }

            let config_dir = get_config_dir()?;
            std::fs::create_dir_all(&config_dir)?;

            let default_config = "# XLEX Configuration File
# See https://github.com/yourname/xlex for documentation

# Default output format: text, json, csv
default_format: text

# Suppress non-essential output
quiet: false

# Enable verbose logging
verbose: false

# Disable colored output
no_color: false

# Default sheet name for new workbooks
default_sheet: Sheet1

# CSV settings
csv_delimiter: \",\"
csv_quote: '\"'

# Date and number format for display
# date_format: YYYY-MM-DD
# number_format: \"#,##0.00\"
";

            std::fs::write(&config_path, default_config)?;

            if !global.quiet {
                println!("{} Created config file: {}", "✓".green(), config_path.display());
            }
        }
        ConfigCommand::Validate => {
            let config_path = get_config_path()?;
            if !config_path.exists() {
                anyhow::bail!("No config file found at {}", config_path.display());
            }

            let content = std::fs::read_to_string(&config_path)?;
            let _yaml: serde_yaml::Value = serde_yaml::from_str(&content)?;

            if !global.quiet {
                println!("{} Config file is valid", "✓".green());
            }
        }
    }

    Ok(())
}

fn run_batch(args: &BatchArgs, global: &GlobalOptions) -> Result<()> {
    use colored::Colorize;
    use std::io::BufRead;

    let reader: Box<dyn std::io::BufRead> = if let Some(ref path) = args.file {
        if path.to_string_lossy() == "-" {
            Box::new(std::io::BufReader::new(std::io::stdin()))
        } else {
            let file = std::fs::File::open(path)?;
            Box::new(std::io::BufReader::new(file))
        }
    } else {
        Box::new(std::io::BufReader::new(std::io::stdin()))
    };

    let mut errors: Vec<(usize, String, String)> = Vec::new();
    let mut success_count = 0;

    for (line_num, line_result) in reader.lines().enumerate() {
        let line = line_result?;
        let line = line.trim();

        // Skip empty lines and comments
        if line.is_empty() || line.starts_with('#') {
            continue;
        }

        if global.verbose {
            println!("{} {}: {}", "[BATCH]".blue(), line_num + 1, line);
        }

        // Parse and execute the command
        // For simplicity, we'll execute using a subprocess
        let output = std::process::Command::new("sh")
            .arg("-c")
            .arg(format!("xlex {}", line))
            .output();

        match output {
            Ok(out) => {
                if out.status.success() {
                    success_count += 1;
                    if !global.quiet && !out.stdout.is_empty() {
                        print!("{}", String::from_utf8_lossy(&out.stdout));
                    }
                } else {
                    let stderr = String::from_utf8_lossy(&out.stderr).to_string();
                    errors.push((line_num + 1, line.to_string(), stderr.clone()));

                    if !args.continue_on_error {
                        anyhow::bail!("Command failed at line {}: {}\n{}", line_num + 1, line, stderr);
                    }
                }
            }
            Err(e) => {
                errors.push((line_num + 1, line.to_string(), e.to_string()));
                if !args.continue_on_error {
                    anyhow::bail!("Failed to execute command at line {}: {}", line_num + 1, e);
                }
            }
        }
    }

    if !global.quiet {
        println!("\n{}: {} commands executed, {} failed",
            "Batch complete".bold(),
            success_count,
            errors.len()
        );

        if !errors.is_empty() {
            println!("\n{}:", "Errors".red());
            for (line, cmd, err) in &errors {
                println!("  Line {}: {}", line, cmd);
                println!("    {}", err.trim().dimmed());
            }
        }
    }

    if !errors.is_empty() && !args.continue_on_error {
        anyhow::bail!("{} commands failed", errors.len());
    }

    Ok(())
}

fn get_alias_path() -> Result<std::path::PathBuf> {
    Ok(get_config_dir()?.join("aliases.yml"))
}

fn run_alias(args: &AliasArgs, global: &GlobalOptions) -> Result<()> {
    use colored::Colorize;

    match &args.command {
        AliasCommand::List => {
            let alias_path = get_alias_path()?;

            if global.format == OutputFormat::Json {
                let aliases: serde_yaml::Mapping = if alias_path.exists() {
                    let content = std::fs::read_to_string(&alias_path)?;
                    serde_yaml::from_str(&content).unwrap_or_default()
                } else {
                    serde_yaml::Mapping::new()
                };
                
                println!("{}", serde_json::to_string_pretty(&aliases)?);
            } else {
                println!("{}:\n", "Aliases".bold());

                // Built-in aliases
                println!("  {} (built-in)", "Built-in".dimmed());
                println!("    {} → {}", "ls".cyan(), "sheet list");
                println!("    {} → {}", "cat".cyan(), "cell get");
                
                // User aliases
                if alias_path.exists() {
                    let content = std::fs::read_to_string(&alias_path)?;
                    let aliases: serde_yaml::Mapping = serde_yaml::from_str(&content).unwrap_or_default();

                    if !aliases.is_empty() {
                        println!("\n  {} (user-defined)", "Custom".dimmed());
                        for (name, cmd) in aliases {
                            if let (serde_yaml::Value::String(n), serde_yaml::Value::String(c)) = (name, cmd) {
                                println!("    {} → {}", n.cyan(), c);
                            }
                        }
                    }
                } else {
                    println!("\n  {}", "(No user-defined aliases)".dimmed());
                }
            }
        }
        AliasCommand::Add { name, command } => {
            if global.dry_run {
                println!("Would add alias: {} → {}", name, command);
                return Ok(());
            }

            let alias_path = get_alias_path()?;
            let config_dir = get_config_dir()?;

            std::fs::create_dir_all(&config_dir)?;

            let mut aliases: serde_yaml::Mapping = if alias_path.exists() {
                let content = std::fs::read_to_string(&alias_path)?;
                serde_yaml::from_str(&content).unwrap_or_default()
            } else {
                serde_yaml::Mapping::new()
            };

            aliases.insert(
                serde_yaml::Value::String(name.clone()),
                serde_yaml::Value::String(command.clone()),
            );

            std::fs::write(&alias_path, serde_yaml::to_string(&aliases)?)?;

            if !global.quiet {
                println!("{} Added alias: {} → {}", "✓".green(), name.cyan(), command);
            }
        }
        AliasCommand::Remove { name } => {
            if global.dry_run {
                println!("Would remove alias: {}", name);
                return Ok(());
            }

            let alias_path = get_alias_path()?;
            if !alias_path.exists() {
                anyhow::bail!("No aliases defined");
            }

            let content = std::fs::read_to_string(&alias_path)?;
            let mut aliases: serde_yaml::Mapping = serde_yaml::from_str(&content).unwrap_or_default();

            let key = serde_yaml::Value::String(name.clone());
            if aliases.remove(&key).is_none() {
                anyhow::bail!("Alias '{}' not found", name);
            }

            std::fs::write(&alias_path, serde_yaml::to_string(&aliases)?)?;

            if !global.quiet {
                println!("{} Removed alias: {}", "✓".green(), name.cyan());
            }
        }
    }

    Ok(())
}

fn run_version(global: &GlobalOptions) -> Result<()> {
    let version = env!("CARGO_PKG_VERSION");
    if global.format == OutputFormat::Json {
        let json = serde_json::json!({
            "version": version,
            "name": "xlex",
        });
        println!("{}", serde_json::to_string_pretty(&json)?);
    } else {
        println!("xlex {}", version);
    }
    Ok(())
}

// Re-export for use in main
pub use clap::CommandFactory;
