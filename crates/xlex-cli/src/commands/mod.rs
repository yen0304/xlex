//! CLI commands for xlex.

mod cell;
mod column;
mod export;
mod formula;
mod import;
mod range;
mod row;
mod search;
pub mod session;
mod sheet;
mod style;
mod template;
mod update;
mod workbook;

use anyhow::Result;
use clap::{Parser, Subcommand};

const LONG_ABOUT: &str = r#"XLEX - A streaming Excel manipulation tool.

XLEX provides CLI-first, streaming-based Excel manipulation for developers
and automation pipelines. It can handle files up to 200MB without memory
exhaustion.

PERFORMANCE TIP:
  For large files (>10MB), use session mode for faster repeated operations:

    $ xlex session <file>

  Session mode loads the file once and keeps it in memory, making subsequent
  commands instant instead of re-parsing the file each time."#;

/// CLI tool for XLEX - A streaming Excel manipulation engine.
#[derive(Parser)]
#[command(name = "xlex")]
#[command(author, version, about, long_about = LONG_ABOUT)]
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

    // Search
    /// Search for text across all sheets (like Ctrl+F in Excel)
    Search(search::SearchArgs),

    // Session management (open → operate → commit)
    /// Open a workbook for editing (creates a session)
    Open(OpenArgs),

    /// Save session changes back to the original file
    Commit,

    /// Discard session changes and close
    Close,

    /// Show current session status
    Status,

    // Utility commands
    /// Generate shell completion scripts
    Completion(CompletionArgs),

    /// Display or modify configuration
    Config(ConfigArgs),

    /// Execute batch commands on a workbook (single open/save cycle)
    Batch(BatchArgs),

    /// Manage command aliases
    Alias(AliasArgs),

    /// Display version information
    Version,

    /// Start interactive mode (REPL)
    Interactive,

    /// Start a session with a pre-loaded workbook (faster for large files)
    #[command(name = "repl")]
    Session(SessionArgs),

    /// Show examples for commands
    Examples(ExamplesArgs),

    /// Generate man page
    Man(ManArgs),

    /// Update xlex to the latest version
    Update(update::UpdateArgs),
}

/// Examples arguments.
#[derive(Parser)]
pub struct ExamplesArgs {
    /// Command to show examples for
    pub command: Option<String>,
    /// Show all examples
    #[arg(long)]
    pub all: bool,
}

/// Man page generation arguments.
#[derive(Parser)]
pub struct ManArgs {
    /// Output directory for man pages
    #[arg(short, long, default_value = ".")]
    pub output_dir: std::path::PathBuf,
    /// Generate for all commands
    #[arg(long)]
    pub all: bool,
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

/// Open arguments.
#[derive(Parser)]
pub struct OpenArgs {
    /// Path to the Excel file to open
    pub file: std::path::PathBuf,
}

/// Batch execution arguments.
#[derive(Parser)]
pub struct BatchArgs {
    /// Path to the Excel file (optional if a session is active)
    pub file: Option<std::path::PathBuf>,

    /// Read commands from a script file (default: stdin)
    #[arg(short = 's', long = "script")]
    pub script: Option<std::path::PathBuf>,

    /// Inline command to execute (can be repeated)
    #[arg(short = 'c', long = "command")]
    pub commands: Vec<String>,

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

/// Session mode arguments.
#[derive(Parser)]
pub struct SessionArgs {
    /// Path to Excel file to load
    pub file: std::path::PathBuf,
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

            // Search
            Commands::Search(args) => search::run(args, &self.global),

            // Session management
            Commands::Open(args) => run_open(args, &self.global),
            Commands::Commit => run_commit(&self.global),
            Commands::Close => run_close(&self.global),
            Commands::Status => run_status(&self.global),

            // Utility
            Commands::Completion(args) => run_completion(args),
            Commands::Config(args) => run_config(args, &self.global),
            Commands::Batch(args) => run_batch(args, &self.global),
            Commands::Alias(args) => run_alias(args, &self.global),
            Commands::Version => run_version(&self.global),
            Commands::Interactive => run_interactive(&self.global),
            Commands::Session(args) => run_session(args, &self.global),
            Commands::Examples(args) => run_examples(args, &self.global),
            Commands::Man(args) => run_man(args, &self.global),
            Commands::Update(args) => update::run(args, &self.global),
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
                    println!(
                        "{}",
                        "(No config file found. Run 'xlex config init' to create one.)".dimmed()
                    );
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
                println!(
                    "{} Created config file: {}",
                    "✓".green(),
                    config_path.display()
                );
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

fn run_open(args: &OpenArgs, global: &GlobalOptions) -> Result<()> {
    use colored::Colorize;

    let state = session::open(&args.file)?;

    if !global.quiet {
        if global.format == OutputFormat::Json {
            let json = serde_json::json!({
                "action": "open",
                "original": state.original_path,
                "working": state.working_path,
            });
            println!("{}", serde_json::to_string_pretty(&json)?);
        } else {
            println!(
                "{} {}",
                "Opened".bold().green(),
                state.original_path.display()
            );
            println!(
                "Working copy: {}",
                state.working_path.display().to_string().dimmed()
            );
            println!();
            println!(
                "Run commands with  {}",
                "xlex batch -c \"cell set Sheet1 A1 hello\"".cyan()
            );
            println!(
                "Or pipe commands   {}",
                "echo 'cell set Sheet1 A1 hello' | xlex batch".cyan()
            );
            println!("Save changes with  {}", "xlex commit".cyan());
            println!("Discard with       {}", "xlex close".cyan());
        }
    }

    Ok(())
}

fn run_commit(global: &GlobalOptions) -> Result<()> {
    use colored::Colorize;

    let state = session::commit()?;

    if !global.quiet {
        if global.format == OutputFormat::Json {
            let json = serde_json::json!({
                "action": "commit",
                "file": state.original_path,
            });
            println!("{}", serde_json::to_string_pretty(&json)?);
        } else {
            println!(
                "{} Changes saved to {}",
                "Committed".bold().green(),
                state.original_path.display()
            );
        }
    }

    Ok(())
}

fn run_close(global: &GlobalOptions) -> Result<()> {
    use colored::Colorize;

    let state = session::close()?;

    if !global.quiet {
        if global.format == OutputFormat::Json {
            let json = serde_json::json!({
                "action": "close",
                "file": state.original_path,
            });
            println!("{}", serde_json::to_string_pretty(&json)?);
        } else {
            println!(
                "{} Changes discarded for {}",
                "Closed".bold().yellow(),
                state.original_path.display()
            );
        }
    }

    Ok(())
}

fn run_status(global: &GlobalOptions) -> Result<()> {
    use colored::Colorize;

    match session::load() {
        Some(state) => {
            if global.format == OutputFormat::Json {
                let json = serde_json::json!({
                    "active": true,
                    "original": state.original_path,
                    "working": state.working_path,
                    "opened_at": state.opened_at,
                });
                println!("{}", serde_json::to_string_pretty(&json)?);
            } else {
                println!("{}", "Session active".bold().green());
                println!("  File:       {}", state.original_path.display());
                println!(
                    "  Working:    {}",
                    state.working_path.display().to_string().dimmed()
                );
                println!("  Opened at:  {}", state.opened_at);
            }
        }
        None => {
            if global.format == OutputFormat::Json {
                let json = serde_json::json!({ "active": false });
                println!("{}", serde_json::to_string_pretty(&json)?);
            } else {
                println!("{}", "No active session".dimmed());
                println!("Run {} to start", "xlex open <file>".cyan());
            }
        }
    }

    Ok(())
}

fn run_batch(args: &BatchArgs, global: &GlobalOptions) -> Result<()> {
    use colored::Colorize;
    use xlex_core::Workbook;

    // Resolve the workbook file: explicit arg > active session
    let file_path = if let Some(ref f) = args.file {
        f.clone()
    } else if let Some(state) = session::load() {
        state.working_path.clone()
    } else {
        anyhow::bail!(
            "No file specified and no active session.\n\
             Usage: xlex batch <file> or xlex open <file> first."
        );
    };

    // Collect commands from: inline -c args, script file, or stdin
    let commands: Vec<String> = if !args.commands.is_empty() {
        args.commands.clone()
    } else {
        use std::io::BufRead;
        let reader: Box<dyn std::io::BufRead> = if let Some(ref path) = args.script {
            let file = std::fs::File::open(path)
                .map_err(|e| anyhow::anyhow!("Cannot open script '{}': {}", path.display(), e))?;
            Box::new(std::io::BufReader::new(file))
        } else {
            Box::new(std::io::BufReader::new(std::io::stdin()))
        };
        reader.lines().collect::<Result<Vec<_>, _>>()?
    };

    // Open workbook ONCE
    let mut workbook = Workbook::open(&file_path)
        .map_err(|e| anyhow::anyhow!("Failed to open '{}': {}", file_path.display(), e))?;

    let mut errors: Vec<(usize, String, String)> = Vec::new();
    let mut success_count = 0;
    let mut modified = false;

    for (line_num, raw_line) in commands.iter().enumerate() {
        let line = raw_line.trim();

        // Skip empty lines and comments
        if line.is_empty() || line.starts_with('#') {
            continue;
        }

        if global.verbose {
            println!("{} {}: {}", "[BATCH]".blue(), line_num + 1, line);
        }

        match execute_batch_command(&mut workbook, line) {
            Ok(msg) => {
                success_count += 1;
                modified = true;
                if !global.quiet && !msg.is_empty() {
                    println!("  {} {}", "✓".green(), msg);
                }
            }
            Err(e) => {
                let err_msg = e.to_string();
                if !global.quiet {
                    eprintln!(
                        "  {} Line {}: {} — {}",
                        "✗".red(),
                        line_num + 1,
                        line,
                        err_msg
                    );
                }
                errors.push((line_num + 1, line.to_string(), err_msg));
                if !args.continue_on_error {
                    anyhow::bail!("Command failed at line {}: {}", line_num + 1, line);
                }
            }
        }
    }

    // Save ONCE if any command succeeded
    if modified {
        workbook
            .save()
            .map_err(|e| anyhow::anyhow!("Failed to save: {}", e))?;
    }

    if !global.quiet {
        println!(
            "\n{}: {} succeeded, {} failed",
            "Batch complete".bold(),
            success_count.to_string().green(),
            errors.len().to_string().red()
        );
    }

    if !errors.is_empty() && !args.continue_on_error {
        anyhow::bail!("{} commands failed", errors.len());
    }

    Ok(())
}

/// Parse and execute a single batch command against an open workbook.
///
/// Supported commands:
///
/// ```text
/// cell set <sheet> <cell> <value>
/// cell clear <sheet> <cell>
/// cell formula <sheet> <cell> <formula>
/// row append <sheet> <val1,val2,...>
/// row insert <sheet> <row_num>
/// row delete <sheet> <row_num>
/// sheet add <name>
/// sheet remove <name>
/// sheet rename <old> <new>
/// ```
fn execute_batch_command(workbook: &mut xlex_core::Workbook, line: &str) -> Result<String> {
    let parts = shell_words(line);
    if parts.is_empty() {
        return Ok(String::new());
    }

    let cmd = parts[0].to_lowercase();
    let sub = parts.get(1).map(|s| s.to_lowercase()).unwrap_or_default();

    match (cmd.as_str(), sub.as_str()) {
        ("cell", "set") => {
            // cell set <sheet> <cell> <value...>
            if parts.len() < 5 {
                anyhow::bail!("Usage: cell set <sheet> <cell> <value>");
            }
            let sheet = &parts[2];
            let cell_str = &parts[3];
            let value_str = parts[4..].join(" ");
            let cell_ref = xlex_core::CellRef::parse(cell_str)?;
            let value = cell::parse_auto_value(&value_str);
            workbook.set_cell(sheet, cell_ref, value)?;
            Ok(format!("cell {} = {}", cell_str, value_str))
        }
        ("cell", "clear") => {
            if parts.len() < 4 {
                anyhow::bail!("Usage: cell clear <sheet> <cell>");
            }
            let sheet = &parts[2];
            let cell_str = &parts[3];
            let cell_ref = xlex_core::CellRef::parse(cell_str)?;
            workbook.clear_cell(sheet, &cell_ref)?;
            Ok(format!("cleared {}", cell_str))
        }
        ("cell", "formula") => {
            if parts.len() < 5 {
                anyhow::bail!("Usage: cell formula <sheet> <cell> <formula>");
            }
            let sheet = &parts[2];
            let cell_str = &parts[3];
            let formula = parts[4..].join(" ");
            let formula = formula.strip_prefix('=').unwrap_or(&formula);
            let cell_ref = xlex_core::CellRef::parse(cell_str)?;
            workbook.set_cell(sheet, cell_ref, xlex_core::CellValue::formula(formula))?;
            Ok(format!("formula {} = ={}", cell_str, formula))
        }
        ("row", "append") => {
            if parts.len() < 4 {
                anyhow::bail!("Usage: row append <sheet> <val1,val2,...>");
            }
            let sheet_name = &parts[2];
            let values_str = parts[3..].join(" ");
            let sheet = workbook
                .get_sheet_mut(sheet_name)
                .ok_or_else(|| anyhow::anyhow!("Sheet '{}' not found", sheet_name))?;

            // Find next row
            let mut max_row: u32 = 0;
            for cell in sheet.cells() {
                if cell.reference.row > max_row {
                    max_row = cell.reference.row;
                }
            }
            let new_row = max_row + 1;

            for (col, val) in values_str.split(',').enumerate() {
                let cell_ref = xlex_core::CellRef::new((col + 1) as u32, new_row);
                let value = cell::parse_auto_value(val.trim());
                sheet.set_cell(cell_ref, value);
            }
            Ok(format!("appended row {}", new_row))
        }
        ("row", "insert") => {
            if parts.len() < 4 {
                anyhow::bail!("Usage: row insert <sheet> <row_num>");
            }
            let sheet_name = &parts[2];
            let row_num: u32 = parts[3]
                .parse()
                .map_err(|_| anyhow::anyhow!("Invalid row number"))?;
            let sheet = workbook
                .get_sheet_mut(sheet_name)
                .ok_or_else(|| anyhow::anyhow!("Sheet '{}' not found", sheet_name))?;
            sheet.insert_rows(row_num, 1);
            Ok(format!("inserted row at {}", row_num))
        }
        ("row", "delete") => {
            if parts.len() < 4 {
                anyhow::bail!("Usage: row delete <sheet> <row_num>");
            }
            let sheet_name = &parts[2];
            let row_num: u32 = parts[3]
                .parse()
                .map_err(|_| anyhow::anyhow!("Invalid row number"))?;
            let sheet = workbook
                .get_sheet_mut(sheet_name)
                .ok_or_else(|| anyhow::anyhow!("Sheet '{}' not found", sheet_name))?;
            sheet.delete_rows(row_num, 1);
            Ok(format!("deleted row {}", row_num))
        }
        ("sheet", "add") => {
            if parts.len() < 3 {
                anyhow::bail!("Usage: sheet add <name>");
            }
            let name = &parts[2];
            workbook.add_sheet(name)?;
            Ok(format!("added sheet '{}'", name))
        }
        ("sheet", "remove") => {
            if parts.len() < 3 {
                anyhow::bail!("Usage: sheet remove <name>");
            }
            let name = &parts[2];
            workbook.remove_sheet(name)?;
            Ok(format!("removed sheet '{}'", name))
        }
        ("sheet", "rename") => {
            if parts.len() < 4 {
                anyhow::bail!("Usage: sheet rename <old> <new>");
            }
            let old = &parts[2];
            let new = &parts[3];
            workbook.rename_sheet(old, new)?;
            Ok(format!("renamed '{}' → '{}'", old, new))
        }
        _ => {
            anyhow::bail!(
                "Unknown batch command: '{} {}'\n\
                 Supported: cell set|clear|formula, row append|insert|delete, sheet add|remove|rename",
                cmd, sub
            );
        }
    }
}

/// Simple shell-like word splitting that respects double quotes.
fn shell_words(input: &str) -> Vec<String> {
    let mut words = Vec::new();
    let mut current = String::new();
    let mut in_quotes = false;
    let chars = input.chars();

    for ch in chars {
        match ch {
            '"' => {
                in_quotes = !in_quotes;
            }
            ' ' | '\t' if !in_quotes => {
                if !current.is_empty() {
                    words.push(std::mem::take(&mut current));
                }
            }
            _ => {
                current.push(ch);
            }
        }
    }
    if !current.is_empty() {
        words.push(current);
    }
    words
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
                    let aliases: serde_yaml::Mapping =
                        serde_yaml::from_str(&content).unwrap_or_default();

                    if !aliases.is_empty() {
                        println!("\n  {} (user-defined)", "Custom".dimmed());
                        for (name, cmd) in aliases {
                            if let (serde_yaml::Value::String(n), serde_yaml::Value::String(c)) =
                                (name, cmd)
                            {
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
            let mut aliases: serde_yaml::Mapping =
                serde_yaml::from_str(&content).unwrap_or_default();

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

fn run_interactive(global: &GlobalOptions) -> Result<()> {
    use colored::Colorize;
    use std::io::{BufRead, Write};

    if !global.quiet {
        println!("{}", "XLEX Interactive Mode".bold().cyan());
        println!(
            "Type {} for help, {} to exit",
            "help".green(),
            "exit".green()
        );
        println!();
    }

    let stdin = std::io::stdin();
    let mut stdout = std::io::stdout();

    loop {
        print!("{} ", "xlex>".bold().green());
        stdout.flush()?;

        let mut line = String::new();
        if stdin.lock().read_line(&mut line)? == 0 {
            break; // EOF
        }

        let line = line.trim();
        if line.is_empty() {
            continue;
        }

        match line.to_lowercase().as_str() {
            "exit" | "quit" | "q" => {
                if !global.quiet {
                    println!("Goodbye!");
                }
                break;
            }
            "help" | "?" => {
                print_interactive_help();
            }
            _ => {
                // Parse and execute command
                let args: Vec<&str> = line.split_whitespace().collect();
                if args.is_empty() {
                    continue;
                }

                // Build the command line with "xlex" prefix
                let mut cmd_args = vec!["xlex"];
                cmd_args.extend(args);

                // Parse and run
                match Cli::try_parse_from(&cmd_args) {
                    Ok(cli) => {
                        if let Err(e) = cli.run() {
                            eprintln!("{}: {}", "error".red(), e);
                        }
                    }
                    Err(e) => {
                        eprintln!("{}", e);
                    }
                }
            }
        }
    }

    Ok(())
}

fn print_interactive_help() {
    use colored::Colorize;

    println!("{}", "Interactive Mode Commands:".bold());
    println!("  {}       - Show this help", "help".cyan());
    println!("  {}       - Exit interactive mode", "exit".cyan());
    println!();
    println!("{}", "XLEX Commands (use without 'xlex' prefix):".bold());
    println!(
        "  {}         - Show workbook information",
        "info <file>".cyan()
    );
    println!("  {}       - List sheets", "sheet list <file>".cyan());
    println!("  {} - Get cell value", "cell get <file> <cell>".cyan());
    println!(
        "  {}   - Set cell value",
        "cell set <file> <cell> <value>".cyan()
    );
    println!();
    println!("{}", "Examples:".bold());
    println!("  info test.xlsx");
    println!("  sheet list test.xlsx");
    println!("  cell get test.xlsx A1");
    println!("  cell set test.xlsx A1 \"Hello World\"");
}

fn run_session(args: &SessionArgs, global: &GlobalOptions) -> Result<()> {
    use colored::Colorize;
    use std::io::{BufRead, Write};
    use std::time::Instant;
    use xlex_core::LazyWorkbook;

    let file_path = &args.file;

    // Load workbook once using lazy loading
    if !global.quiet {
        println!("{} {}...", "Loading".bold().cyan(), file_path.display());
    }

    let start = Instant::now();
    let workbook = LazyWorkbook::open(file_path)
        .map_err(|e| anyhow::anyhow!("Failed to open workbook: {}", e))?;
    let load_time = start.elapsed();

    if !global.quiet {
        println!(
            "{} in {:.2}s",
            "Loaded".bold().green(),
            load_time.as_secs_f64()
        );
        println!();
        println!("{}", "Session Mode".bold().cyan());
        println!(
            "Type {} for help, {} to exit",
            "help".green(),
            "exit".green()
        );
        println!();
    }

    let stdin = std::io::stdin();
    let mut stdout = std::io::stdout();

    loop {
        print!("{} ", "session>".bold().yellow());
        stdout.flush()?;

        let mut line = String::new();
        if stdin.lock().read_line(&mut line)? == 0 {
            break; // EOF
        }

        let line = line.trim();
        if line.is_empty() {
            continue;
        }

        let parts: Vec<&str> = line.split_whitespace().collect();
        if parts.is_empty() {
            continue;
        }

        let cmd = parts[0].to_lowercase();
        let args = &parts[1..];

        match cmd.as_str() {
            "exit" | "quit" | "q" => {
                if !global.quiet {
                    println!("Goodbye!");
                }
                break;
            }
            "help" | "?" => {
                print_session_help();
            }
            "info" => {
                run_session_info(&workbook, global);
            }
            "sheets" | "sheet" => {
                if args.first().copied() == Some("list") || args.is_empty() {
                    run_session_sheets(&workbook, global);
                } else {
                    eprintln!("{}: unknown sheet subcommand", "error".red());
                    eprintln!("Use: sheets, sheet list");
                }
            }
            "cell" => {
                if args.len() < 2 {
                    eprintln!("{}: usage: cell <sheet> <ref>", "error".red());
                    eprintln!("Example: cell Sheet1 A1");
                } else {
                    run_session_cell(&workbook, args[0], args[1], global);
                }
            }
            "row" => {
                if args.len() < 2 {
                    eprintln!("{}: usage: row <sheet> <number>", "error".red());
                    eprintln!("Example: row Sheet1 1");
                } else {
                    match args[1].parse::<u32>() {
                        Ok(row_num) => run_session_row(&workbook, args[0], row_num, global),
                        Err(_) => eprintln!("{}: invalid row number", "error".red()),
                    }
                }
            }
            "search" | "find" => {
                if args.is_empty() {
                    eprintln!("{}: usage: search <pattern> [sheet]", "error".red());
                    eprintln!("Example: search revenue");
                    eprintln!("Example: search error Sheet1");
                } else {
                    let pattern = args[0];
                    let sheet_filter = args.get(1).copied();
                    run_session_search(&workbook, pattern, sheet_filter, global);
                }
            }
            _ => {
                eprintln!("{}: unknown command '{}'", "error".red(), cmd);
                eprintln!("Type 'help' for available commands");
            }
        }
    }

    Ok(())
}

fn print_session_help() {
    use colored::Colorize;

    println!("{}", "Session Mode Commands:".bold());
    println!("  {}       - Show this help", "help".cyan());
    println!("  {}       - Exit session mode", "exit".cyan());
    println!();
    println!("{}", "Workbook Commands:".bold());
    println!("  {}           - Show workbook information", "info".cyan());
    println!("  {}         - List all sheets", "sheets".cyan());
    println!("  {}  - Get cell value", "cell <sheet> <ref>".cyan());
    println!("  {} - Get row values", "row <sheet> <number>".cyan());
    println!(
        "  {} - Search across all sheets",
        "search <pattern> [sheet]".cyan()
    );
    println!();
    println!("{}", "Examples:".bold());
    println!("  info");
    println!("  sheets");
    println!("  cell Sheet1 A1");
    println!("  cell Sheet1 B2:D5");
    println!("  row Sheet1 1");
    println!("  search revenue");
    println!("  search error Sheet1");
}

fn run_session_info(workbook: &xlex_core::LazyWorkbook, global: &GlobalOptions) {
    use colored::Colorize;

    let sheets = workbook.sheet_names();

    if global.format == OutputFormat::Json {
        let json = serde_json::json!({
            "sheet_count": sheets.len(),
            "sheets": sheets,
        });
        println!("{}", serde_json::to_string_pretty(&json).unwrap());
    } else {
        println!("{}: {}", "Sheet count".bold(), sheets.len());
        println!("{}: {}", "Sheets".bold(), sheets.join(", "));
    }
}

fn run_session_sheets(workbook: &xlex_core::LazyWorkbook, global: &GlobalOptions) {
    use colored::Colorize;

    let sheets = workbook.sheet_names();

    if global.format == OutputFormat::Json {
        let json = serde_json::json!({
            "sheets": sheets,
        });
        println!("{}", serde_json::to_string_pretty(&json).unwrap());
    } else {
        println!("{}", "Sheets:".bold());
        for (i, name) in sheets.iter().enumerate() {
            println!("  {}. {}", i + 1, name.cyan());
        }
    }
}

fn run_session_cell(
    workbook: &xlex_core::LazyWorkbook,
    sheet_name: &str,
    cell_ref_str: &str,
    global: &GlobalOptions,
) {
    use colored::Colorize;
    use xlex_core::CellRef;

    let cell_ref = match CellRef::parse(cell_ref_str) {
        Ok(r) => r,
        Err(e) => {
            eprintln!(
                "{}: invalid cell reference '{}': {}",
                "error".red(),
                cell_ref_str,
                e
            );
            return;
        }
    };

    match workbook.read_cell(sheet_name, &cell_ref) {
        Ok(Some(value)) => {
            if global.format == OutputFormat::Json {
                let json = serde_json::json!({
                    "sheet": sheet_name,
                    "cell": cell_ref_str,
                    "value": value.to_string(),
                });
                println!("{}", serde_json::to_string_pretty(&json).unwrap());
            } else {
                println!("{}", value);
            }
        }
        Ok(None) => {
            if global.format == OutputFormat::Json {
                let json = serde_json::json!({
                    "sheet": sheet_name,
                    "cell": cell_ref_str,
                    "value": null,
                });
                println!("{}", serde_json::to_string_pretty(&json).unwrap());
            } else {
                println!("(empty)");
            }
        }
        Err(e) => {
            eprintln!("{}: {}", "error".red(), e);
        }
    }
}

fn run_session_row(
    workbook: &xlex_core::LazyWorkbook,
    sheet_name: &str,
    row_num: u32,
    global: &GlobalOptions,
) {
    use colored::Colorize;

    // Stream rows and find the one we want
    match workbook.stream_rows(sheet_name) {
        Ok(rows) => {
            let mut found = false;
            for row in rows {
                if row.row_number == row_num {
                    found = true;
                    if global.format == OutputFormat::Json {
                        let cells_map: std::collections::HashMap<String, String> = row
                            .cells
                            .iter()
                            .map(|(cell_ref, value)| (cell_ref.to_string(), value.to_string()))
                            .collect();
                        let json = serde_json::json!({
                            "sheet": sheet_name,
                            "row": row_num,
                            "cells": cells_map,
                        });
                        println!("{}", serde_json::to_string_pretty(&json).unwrap());
                    } else {
                        let values: Vec<String> =
                            row.cells.iter().map(|(_, v)| v.to_string()).collect();
                        println!("{}", values.join("\t"));
                    }
                    break;
                }
            }
            if !found {
                eprintln!("{}: row {} not found", "error".red(), row_num);
            }
        }
        Err(e) => {
            eprintln!("{}: {}", "error".red(), e);
        }
    }
}

fn run_session_search(
    workbook: &xlex_core::LazyWorkbook,
    pattern: &str,
    sheet_filter: Option<&str>,
    global: &GlobalOptions,
) {
    use colored::Colorize;

    let sheets = workbook.sheet_names();
    let search_sheets: Vec<&str> = if let Some(name) = sheet_filter {
        if !workbook.has_sheet(name) {
            eprintln!("{}: sheet '{}' not found", "error".red(), name);
            return;
        }
        vec![name]
    } else {
        sheets.iter().map(|s| s.as_str()).collect()
    };

    let pattern_lower = pattern.to_lowercase();

    #[derive(serde::Serialize)]
    struct Match {
        sheet: String,
        cell: String,
        value: String,
    }

    let mut matches: Vec<Match> = Vec::new();

    for sheet_name in &search_sheets {
        match workbook.stream_rows(sheet_name) {
            Ok(rows) => {
                for row in rows {
                    for (cell_ref, cell_value) in &row.cells {
                        let value = cell_value.to_string();
                        if value.to_lowercase().contains(&pattern_lower) {
                            matches.push(Match {
                                sheet: sheet_name.to_string(),
                                cell: cell_ref.to_a1(),
                                value,
                            });
                        }
                    }
                }
            }
            Err(e) => {
                eprintln!(
                    "{}: failed to read sheet '{}': {}",
                    "error".red(),
                    sheet_name,
                    e
                );
            }
        }
    }

    if global.format == OutputFormat::Json {
        let json = serde_json::json!({
            "pattern": pattern,
            "count": matches.len(),
            "matches": matches,
        });
        println!("{}", serde_json::to_string_pretty(&json).unwrap());
    } else if matches.is_empty() {
        println!(
            "{}: no matches found for \"{}\"",
            "search".yellow(),
            pattern
        );
    } else {
        println!(
            "Found {} match{} for \"{}\":\n",
            matches.len().to_string().green(),
            if matches.len() == 1 { "" } else { "es" },
            pattern.cyan()
        );
        let mut current_sheet = "";
        for m in &matches {
            if m.sheet != current_sheet {
                if !current_sheet.is_empty() {
                    println!();
                }
                println!("  {} {}", "Sheet:".bold(), m.sheet.bold());
                current_sheet = &m.sheet;
            }
            println!("    {} = {}", m.cell.yellow(), m.value.dimmed());
        }
        println!();
    }
}

fn run_examples(args: &ExamplesArgs, _global: &GlobalOptions) -> Result<()> {
    if args.all {
        print_all_examples();
        return Ok(());
    }

    if let Some(command) = &args.command {
        print_command_help_with_examples(command);
    } else {
        print_overview_help();
    }

    Ok(())
}

fn print_overview_help() {
    use colored::Colorize;

    println!(
        "{}",
        "XLEX - A streaming Excel manipulation tool".bold().cyan()
    );
    println!();
    println!("{}", "USAGE:".bold());
    println!("    xlex <COMMAND> [OPTIONS]");
    println!();
    println!("{}", "COMMANDS:".bold());
    println!("    {}      Show workbook information", "info".green());
    println!("    {}     Create a new workbook", "create".green());
    println!("    {}     Sheet operations", "sheet".green());
    println!("    {}      Cell operations", "cell".green());
    println!("    {}       Row operations", "row".green());
    println!("    {}    Column operations", "column".green());
    println!("    {}     Range operations", "range".green());
    println!("    {}     Style operations", "style".green());
    println!("    {}   Formula operations", "formula".green());
    println!("    {}  Template operations", "template".green());
    println!("    {}    Import from external format", "import".green());
    println!("    {}    Export to external format", "export".green());
    println!("    {}   Convert between formats", "convert".green());
    println!();
    println!("{}", "QUICK EXAMPLES:".bold());
    println!("    xlex info workbook.xlsx");
    println!("    xlex create new.xlsx");
    println!("    xlex cell get workbook.xlsx A1");
    println!("    xlex cell set workbook.xlsx A1 \"Hello\"");
    println!("    xlex sheet list workbook.xlsx");
    println!("    xlex export csv workbook.xlsx data.csv");
    println!();
    println!(
        "Run {} for examples for a specific command",
        "xlex examples <command>".yellow()
    );
    println!("Run {} for all examples", "xlex examples --all".yellow());
}

fn print_all_examples() {
    use colored::Colorize;

    println!("{}", "XLEX Command Examples".bold().cyan());
    println!();

    // Workbook examples
    println!("{}", "WORKBOOK OPERATIONS:".bold());
    println!("  # Show workbook information");
    println!("  {} workbook.xlsx", "xlex info".green());
    println!("  {} workbook.xlsx --format json", "xlex info".green());
    println!();
    println!("  # Create a new workbook");
    println!("  {} new.xlsx", "xlex create".green());
    println!(
        "  {} new.xlsx --sheets Sales,Inventory,Summary",
        "xlex create".green()
    );
    println!();
    println!("  # Clone a workbook");
    println!("  {} original.xlsx copy.xlsx", "xlex clone".green());
    println!();

    // Sheet examples
    println!("{}", "SHEET OPERATIONS:".bold());
    println!("  # List sheets");
    println!("  {} workbook.xlsx", "xlex sheet list".green());
    println!();
    println!("  # Add a new sheet");
    println!("  {} workbook.xlsx NewSheet", "xlex sheet add".green());
    println!();
    println!("  # Rename a sheet");
    println!(
        "  {} workbook.xlsx OldName NewName",
        "xlex sheet rename".green()
    );
    println!();
    println!("  # Remove a sheet");
    println!(
        "  {} workbook.xlsx SheetToRemove",
        "xlex sheet remove".green()
    );
    println!();

    // Cell examples
    println!("{}", "CELL OPERATIONS:".bold());
    println!("  # Get a cell value");
    println!("  {} workbook.xlsx A1", "xlex cell get".green());
    println!("  {} workbook.xlsx B2 -s Sales", "xlex cell get".green());
    println!();
    println!("  # Set a cell value");
    println!(
        "  {} workbook.xlsx A1 \"Hello World\"",
        "xlex cell set".green()
    );
    println!(
        "  {} workbook.xlsx B2 123.45 -s Sales",
        "xlex cell set".green()
    );
    println!();
    println!("  # Set a formula");
    println!(
        "  {} workbook.xlsx A5 \"=SUM(A1:A4)\"",
        "xlex cell formula".green()
    );
    println!();
    println!("  # Batch update cells");
    println!(
        "  echo 'A1=Hello' | {} workbook.xlsx",
        "xlex cell batch".green()
    );
    println!();

    // Row/Column examples
    println!("{}", "ROW & COLUMN OPERATIONS:".bold());
    println!("  # Get a row");
    println!("  {} workbook.xlsx 1", "xlex row get".green());
    println!();
    println!("  # Append a row");
    println!(
        "  {} workbook.xlsx Value1,Value2,Value3",
        "xlex row append".green()
    );
    println!();
    println!("  # Get a column");
    println!("  {} workbook.xlsx A", "xlex column get".green());
    println!();
    println!("  # Set column width");
    println!("  {} workbook.xlsx A 20", "xlex column width".green());
    println!();

    // Range examples
    println!("{}", "RANGE OPERATIONS:".bold());
    println!("  # Get a range");
    println!("  {} workbook.xlsx A1:C10", "xlex range get".green());
    println!();
    println!("  # Copy a range");
    println!("  {} workbook.xlsx A1:C10 E1", "xlex range copy".green());
    println!();
    println!("  # Merge cells");
    println!("  {} workbook.xlsx A1:D1", "xlex range merge".green());
    println!();

    // Import/Export examples
    println!("{}", "IMPORT/EXPORT:".bold());
    println!("  # Export to CSV");
    println!("  {} workbook.xlsx output.csv", "xlex export csv".green());
    println!();
    println!("  # Export to JSON");
    println!("  {} workbook.xlsx output.json", "xlex export json".green());
    println!();
    println!("  # Import from CSV");
    println!("  {} data.csv workbook.xlsx", "xlex import csv".green());
    println!();
    println!("  # Convert between formats");
    println!("  {} data.csv output.xlsx", "xlex convert".green());
    println!("  {} workbook.xlsx output.json", "xlex convert".green());
    println!();

    // Template examples
    println!("{}", "TEMPLATE OPERATIONS:".bold());
    println!("  # Create a template");
    println!(
        "  {} template.xlsx --type invoice",
        "xlex template init".green()
    );
    println!();
    println!("  # List placeholders");
    println!("  {} template.xlsx", "xlex template list".green());
    println!();
    println!("  # Apply template with variables");
    println!(
        "  {} template.xlsx output.xlsx -v vars.json",
        "xlex template apply".green()
    );
    println!(
        "  {} template.xlsx output.xlsx -D name=John -D date=2024-01-01",
        "xlex template apply".green()
    );
    println!();
    println!("  # Batch template processing");
    println!(
        "  {} template.xlsx output.xlsx -v data.json --per-record",
        "xlex template apply".green()
    );
    println!();

    // Formula examples
    println!("{}", "FORMULA OPERATIONS:".bold());
    println!("  # List all formulas");
    println!("  {} workbook.xlsx", "xlex formula list".green());
    println!();
    println!("  # Validate formulas");
    println!("  {} workbook.xlsx", "xlex formula validate".green());
    println!();
    println!("  # Get formula statistics");
    println!("  {} workbook.xlsx", "xlex formula stats".green());
}

fn print_command_help_with_examples(command: &str) {
    use colored::Colorize;

    match command.to_lowercase().as_str() {
        "info" => {
            println!("{}", "xlex info - Display workbook information".bold());
            println!();
            println!("{}", "USAGE:".bold());
            println!("    xlex info <FILE> [OPTIONS]");
            println!();
            println!("{}", "EXAMPLES:".bold());
            println!("    xlex info workbook.xlsx");
            println!("    xlex info workbook.xlsx --format json");
            println!("    xlex info workbook.xlsx -v");
        }
        "create" => {
            println!("{}", "xlex create - Create a new workbook".bold());
            println!();
            println!("{}", "USAGE:".bold());
            println!("    xlex create <FILE> [OPTIONS]");
            println!();
            println!("{}", "OPTIONS:".bold());
            println!("    --sheets <NAMES>    Comma-separated sheet names");
            println!("    --force             Overwrite existing file");
            println!();
            println!("{}", "EXAMPLES:".bold());
            println!("    xlex create new.xlsx");
            println!("    xlex create report.xlsx --sheets Summary,Data,Charts");
            println!("    xlex create backup.xlsx --force");
        }
        "sheet" => {
            println!("{}", "xlex sheet - Sheet operations".bold());
            println!();
            println!("{}", "SUBCOMMANDS:".bold());
            println!("    list    List all sheets");
            println!("    add     Add a new sheet");
            println!("    remove  Remove a sheet");
            println!("    rename  Rename a sheet");
            println!("    copy    Copy a sheet");
            println!("    move    Move a sheet");
            println!("    info    Show sheet information");
            println!();
            println!("{}", "EXAMPLES:".bold());
            println!("    xlex sheet list workbook.xlsx");
            println!("    xlex sheet add workbook.xlsx NewSheet");
            println!("    xlex sheet rename workbook.xlsx OldName NewName");
            println!("    xlex sheet copy workbook.xlsx Sheet1 Sheet1_Copy");
            println!("    xlex sheet move workbook.xlsx Sheet1 2");
        }
        "cell" => {
            println!("{}", "xlex cell - Cell operations".bold());
            println!();
            println!("{}", "SUBCOMMANDS:".bold());
            println!("    get      Get cell value");
            println!("    set      Set cell value");
            println!("    formula  Set cell formula");
            println!("    clear    Clear cell content");
            println!("    batch    Batch update cells");
            println!();
            println!("{}", "EXAMPLES:".bold());
            println!("    xlex cell get workbook.xlsx A1");
            println!("    xlex cell get workbook.xlsx B2 -s Sales");
            println!("    xlex cell set workbook.xlsx A1 \"Hello World\"");
            println!("    xlex cell formula workbook.xlsx C10 \"=SUM(C1:C9)\"");
            println!("    echo 'A1=Hello\\nB1=World' | xlex cell batch workbook.xlsx");
        }
        "template" => {
            println!("{}", "xlex template - Template operations".bold());
            println!();
            println!("{}", "SUBCOMMANDS:".bold());
            println!("    init      Create a new template");
            println!("    list      List placeholders");
            println!("    validate  Validate template");
            println!("    apply     Apply template with variables");
            println!("    preview   Preview template rendering");
            println!();
            println!("{}", "TEMPLATE FEATURES:".bold());
            println!("    {{{{name}}}}                  Simple placeholder");
            println!("    {{{{name|upper}}}}            Filter (upper, lower, currency, etc.)");
            println!("    {{{{#if condition}}}}...{{{{/if}}}}  Conditional");
            println!("    {{{{#row-repeat items}}}}    Row repetition");
            println!();
            println!("{}", "EXAMPLES:".bold());
            println!("    xlex template init report.xlsx --type invoice");
            println!("    xlex template list template.xlsx");
            println!("    xlex template apply template.xlsx output.xlsx -v vars.json");
            println!("    xlex template apply template.xlsx output.xlsx -D name=John");
            println!("    xlex template apply template.xlsx output.xlsx --per-record -v data.json");
        }
        _ => {
            println!("No detailed help available for '{}'", command);
            println!("Run {} for general help", "xlex --help".yellow());
        }
    }
}

fn run_man(args: &ManArgs, global: &GlobalOptions) -> Result<()> {
    use std::io::Write;

    if global.dry_run {
        println!("Would generate man pages in {}", args.output_dir.display());
        return Ok(());
    }

    std::fs::create_dir_all(&args.output_dir)?;

    let man_content = generate_man_page();
    let man_path = args.output_dir.join("xlex.1");

    let mut file = std::fs::File::create(&man_path)?;
    file.write_all(man_content.as_bytes())?;

    if !global.quiet {
        println!("Generated man page: {}", man_path.display());
    }

    Ok(())
}

fn generate_man_page() -> String {
    let version = env!("CARGO_PKG_VERSION");

    format!(
        r#".TH XLEX 1 "2024" "xlex {version}" "User Commands"
.SH NAME
xlex \- A streaming Excel manipulation tool
.SH SYNOPSIS
.B xlex
.I COMMAND
.RI [ OPTIONS ]
.SH DESCRIPTION
XLEX is a CLI-first, streaming-based Excel manipulation tool for developers
and automation pipelines. It can handle files up to 200MB without memory exhaustion.
.SH COMMANDS
.TP
.B info \fIFILE\fR
Display workbook information
.TP
.B create \fIFILE\fR
Create a new workbook
.TP
.B clone \fISOURCE\fR \fIDEST\fR
Clone a workbook
.TP
.B sheet \fISUBCOMMAND\fR
Sheet operations (list, add, remove, rename, copy, move)
.TP
.B cell \fISUBCOMMAND\fR
Cell operations (get, set, formula, clear, batch)
.TP
.B row \fISUBCOMMAND\fR
Row operations (get, append, insert, delete)
.TP
.B column \fISUBCOMMAND\fR
Column operations (get, insert, delete, width)
.TP
.B range \fISUBCOMMAND\fR
Range operations (get, copy, move, merge, sort, filter)
.TP
.B style \fISUBCOMMAND\fR
Style operations (list, get, apply)
.TP
.B formula \fISUBCOMMAND\fR
Formula operations (list, validate, stats, refs)
.TP
.B template \fISUBCOMMAND\fR
Template operations (init, list, validate, apply, preview)
.TP
.B import \fIFORMAT\fR \fISOURCE\fR \fIDEST\fR
Import from external format (csv, json, ndjson)
.TP
.B export \fIFORMAT\fR \fISOURCE\fR \fIDEST\fR
Export to external format (csv, json, ndjson)
.TP
.B convert \fIINPUT\fR \fIOUTPUT\fR
Convert between formats
.SH GLOBAL OPTIONS
.TP
.B \-q, \-\-quiet
Suppress all output except errors
.TP
.B \-v, \-\-verbose
Enable verbose output
.TP
.B \-f, \-\-format \fIFORMAT\fR
Output format (text, json, csv, ndjson)
.TP
.B \-\-no\-color
Disable colored output
.TP
.B \-\-dry\-run
Perform a dry run without making changes
.TP
.B \-o, \-\-output \fIFILE\fR
Write output to file instead of stdout
.SH EXAMPLES
.TP
Show workbook information:
.B xlex info workbook.xlsx
.TP
Create a new workbook:
.B xlex create new.xlsx \-\-sheets Sales,Inventory
.TP
Get a cell value:
.B xlex cell get workbook.xlsx A1
.TP
Set a cell value:
.B xlex cell set workbook.xlsx A1 "Hello World"
.TP
Export to CSV:
.B xlex export csv workbook.xlsx data.csv
.TP
Apply a template:
.B xlex template apply template.xlsx output.xlsx \-v vars.json
.SH EXIT STATUS
.TP
.B 0
Success
.TP
.B 1
General error
.TP
.B 2
Invalid arguments
.TP
.B 10\-19
File errors (not found, permission denied, etc.)
.TP
.B 20\-29
Parse errors (invalid xlsx, corrupt file, etc.)
.TP
.B 30\-39
Validation errors
.SH ENVIRONMENT
.TP
.B XLEX_CONFIG
Path to configuration file
.TP
.B XLEX_NO_COLOR
Disable colored output when set to 1
.TP
.B XLEX_LOG_FILE
Path to log file for error logging
.SH FILES
.TP
.I ~/.config/xlex/config.yml
User configuration file
.TP
.I .xlex.yml
Project configuration file
.SH SEE ALSO
.BR xlsx (5)
.SH BUGS
Report bugs at: https://github.com/xlex/xlex/issues
.SH AUTHORS
XLEX Contributors
"#,
        version = version
    )
}

// Re-export for use in main
pub use clap::CommandFactory;
