# CLI Interface Specification

## ADDED Requirements

### Requirement: Command Structure

The system SHALL provide a hierarchical command structure following `xlex <resource> <action> [args]`.

#### Scenario: Top-level help
- **GIVEN** the command `xlex --help`
- **WHEN** executed
- **THEN** output SHALL list all available subcommands with descriptions

#### Scenario: Subcommand help
- **GIVEN** the command `xlex sheet --help`
- **WHEN** executed
- **THEN** output SHALL list all sheet actions (list, add, remove, etc.)

#### Scenario: Action help
- **GIVEN** the command `xlex sheet add --help`
- **WHEN** executed
- **THEN** output SHALL show usage, arguments, and options for sheet add

#### Scenario: Version display
- **GIVEN** the command `xlex --version`
- **WHEN** executed
- **THEN** output SHALL show version number (e.g., "xlex 1.0.0")

### Requirement: Global Flags

The system SHALL support global flags applicable to all commands.

#### Scenario: Quiet mode
- **GIVEN** the command `xlex --quiet sheet add report.xlsx "New"`
- **WHEN** executed
- **THEN** only errors SHALL be output
- **AND** success messages SHALL be suppressed

#### Scenario: Verbose mode
- **GIVEN** the command `xlex --verbose sheet add report.xlsx "New"`
- **WHEN** executed
- **THEN** detailed operation logs SHALL be output

#### Scenario: Output format
- **GIVEN** the command `xlex --format json sheet list report.xlsx`
- **WHEN** executed
- **THEN** output SHALL be JSON format

#### Scenario: Output file
- **GIVEN** the command `xlex cell get report.xlsx "Data" A1 --output result.txt`
- **WHEN** executed
- **THEN** result SHALL be written to result.txt

#### Scenario: No color
- **GIVEN** the command `xlex --no-color info report.xlsx`
- **WHEN** executed
- **THEN** output SHALL not contain ANSI color codes

#### Scenario: Force color
- **GIVEN** the command `xlex --color info report.xlsx`
- **WHEN** piped to another command
- **THEN** output SHALL contain ANSI color codes

### Requirement: Input/Output Handling

The system SHALL support flexible input/output handling.

#### Scenario: Read from stdin
- **GIVEN** the command `cat data.csv | xlex from csv report.xlsx "Data"`
- **WHEN** executed
- **THEN** stdin SHALL be read as input

#### Scenario: Write to stdout
- **GIVEN** the command `xlex to csv report.xlsx "Data"`
- **WHEN** executed without --output
- **THEN** output SHALL go to stdout

#### Scenario: Output to file
- **GIVEN** the command `xlex to csv report.xlsx "Data" --output data.csv`
- **WHEN** executed
- **THEN** output SHALL be written to data.csv

#### Scenario: In-place modification
- **GIVEN** the command `xlex cell set report.xlsx "Data" A1 "New"`
- **WHEN** executed without --output
- **THEN** report.xlsx SHALL be modified in place

#### Scenario: Output to different file
- **GIVEN** the command `xlex cell set report.xlsx "Data" A1 "New" --output modified.xlsx`
- **WHEN** executed
- **THEN** modified.xlsx SHALL contain the change
- **AND** report.xlsx SHALL remain unchanged

#### Scenario: Dry run
- **GIVEN** the command `xlex --dry-run cell set report.xlsx "Data" A1 "New"`
- **WHEN** executed
- **THEN** the operation SHALL be validated but not executed
- **AND** output SHALL describe what would happen

### Requirement: Shell Completion

The system SHALL generate shell completion scripts via `xlex completion <shell>`.

#### Scenario: Bash completion
- **GIVEN** the command `xlex completion bash`
- **WHEN** executed
- **THEN** output SHALL be valid bash completion script

#### Scenario: Zsh completion
- **GIVEN** the command `xlex completion zsh`
- **WHEN** executed
- **THEN** output SHALL be valid zsh completion script

#### Scenario: Fish completion
- **GIVEN** the command `xlex completion fish`
- **WHEN** executed
- **THEN** output SHALL be valid fish completion script

#### Scenario: PowerShell completion
- **GIVEN** the command `xlex completion powershell`
- **WHEN** executed
- **THEN** output SHALL be valid PowerShell completion script

### Requirement: Configuration

The system SHALL support configuration via `xlex config`.

#### Scenario: Show config
- **GIVEN** the command `xlex config show`
- **WHEN** executed
- **THEN** output SHALL show current configuration

#### Scenario: Set config value
- **GIVEN** the command `xlex config set default-format json`
- **WHEN** executed
- **THEN** default-format SHALL be set to json

#### Scenario: Get config value
- **GIVEN** the command `xlex config get default-format`
- **WHEN** executed
- **THEN** output SHALL be the current value

#### Scenario: Reset config
- **GIVEN** the command `xlex config reset`
- **WHEN** executed
- **THEN** all config SHALL be reset to defaults

#### Scenario: Config file location
- **GIVEN** the environment variable XLEX_CONFIG set
- **WHEN** xlex runs
- **THEN** config SHALL be read from the specified path

### Requirement: Project Configuration File

The system SHALL support project-level configuration via `.xlex.yml` or `.xlex.yaml`.

#### Scenario: Load project config
- **GIVEN** a `.xlex.yml` file in the current directory
- **WHEN** xlex runs
- **THEN** settings from `.xlex.yml` SHALL be applied

#### Scenario: Config file search
- **GIVEN** no `.xlex.yml` in current directory
- **WHEN** xlex runs
- **THEN** it SHALL search parent directories up to root for `.xlex.yml`

#### Scenario: Config file format
- **GIVEN** a `.xlex.yml` file
- **WHEN** parsed
- **THEN** it SHALL support YAML format with the following structure:
  ```yaml
  # .xlex.yml
  default_format: json
  string_cache_size: 20000
  no_color: false
  quiet: false
  ```

#### Scenario: Override with CLI flags
- **GIVEN** `.xlex.yml` with `default_format: json`
- **WHEN** executing `xlex --format csv sheet list report.xlsx`
- **THEN** CLI flag SHALL override config file setting

#### Scenario: Override with environment
- **GIVEN** `.xlex.yml` with `default_format: json`
- **AND** XLEX_DEFAULT_FORMAT=csv
- **WHEN** executing `xlex sheet list report.xlsx`
- **THEN** environment variable SHALL override config file setting

#### Scenario: Config precedence
- **GIVEN** configuration from multiple sources
- **WHEN** xlex runs
- **THEN** precedence SHALL be (highest to lowest):
  1. CLI flags
  2. Environment variables
  3. Project config (`.xlex.yml`)
  4. User config (`~/.config/xlex/config.yml`)
  5. Default values

#### Scenario: Available config options
- **GIVEN** a `.xlex.yml` file
- **WHEN** configuring
- **THEN** the following options SHALL be supported:
  - `default_format`: Output format (text, json, csv)
  - `string_cache_size`: SharedStrings LRU cache size
  - `no_color`: Disable colored output
  - `quiet`: Suppress non-essential output
  - `verbose`: Enable verbose logging
  - `default_sheet`: Default sheet name for new workbooks
  - `date_format`: Default date format for display
  - `number_format`: Default number format for display
  - `csv_delimiter`: Default CSV delimiter
  - `csv_quote`: Default CSV quote character

#### Scenario: Init config
- **GIVEN** the command `xlex config init`
- **WHEN** executed
- **THEN** a `.xlex.yml` file SHALL be created with default values and comments

#### Scenario: Validate config
- **GIVEN** the command `xlex config validate`
- **WHEN** executed
- **THEN** the config file SHALL be validated
- **AND** errors SHALL be reported if invalid

#### Scenario: Show effective config
- **GIVEN** the command `xlex config show --effective`
- **WHEN** executed
- **THEN** output SHALL show merged config from all sources
- **AND** indicate the source of each value

#### Scenario: JSON config format
- **GIVEN** a `.xlex.json` file in the current directory
- **WHEN** xlex runs
- **THEN** JSON format config SHALL also be supported

#### Scenario: Ignore config file
- **GIVEN** the command `xlex --no-config sheet list report.xlsx`
- **WHEN** executed
- **THEN** project config file SHALL be ignored

### Requirement: Progress Indication

The system SHALL provide progress indication for long operations.

#### Scenario: Progress bar
- **GIVEN** a long operation (e.g., appending 100k rows)
- **WHEN** stdout is a TTY
- **THEN** a progress bar SHALL be displayed

#### Scenario: No progress in pipe
- **GIVEN** a long operation
- **WHEN** stdout is piped
- **THEN** no progress indication SHALL be output
- **AND** only final result SHALL be output

#### Scenario: Disable progress
- **GIVEN** the command `xlex --no-progress row append report.xlsx "Data" --from large.csv`
- **WHEN** executed
- **THEN** no progress indication SHALL be shown

### Requirement: Batch Operations

The system SHALL support batch operations via `xlex batch`.

#### Scenario: Batch from file
- **GIVEN** a file commands.txt with one command per line
- **WHEN** executing `xlex batch --file commands.txt`
- **THEN** all commands SHALL be executed sequentially

#### Scenario: Batch from stdin
- **GIVEN** commands piped to stdin
- **WHEN** executing `xlex batch`
- **THEN** each line SHALL be executed as a command

#### Scenario: Batch with continue on error
- **GIVEN** the command `xlex batch --file commands.txt --continue-on-error`
- **WHEN** one command fails
- **THEN** subsequent commands SHALL still execute
- **AND** errors SHALL be reported at the end

#### Scenario: Batch transaction
- **GIVEN** the command `xlex batch --file commands.txt --transaction`
- **WHEN** one command fails
- **THEN** all changes SHALL be rolled back

### Requirement: Alias Support

The system SHALL support command aliases.

#### Scenario: Built-in aliases
- **GIVEN** the command `xlex ls report.xlsx`
- **WHEN** executed
- **THEN** it SHALL be equivalent to `xlex sheet list report.xlsx`

#### Scenario: Create alias
- **GIVEN** the command `xlex alias add headers "range style --bold --bg-color #4472C4"`
- **WHEN** executed
- **THEN** the alias "headers" SHALL be created

#### Scenario: Use alias
- **GIVEN** an alias "headers" defined
- **WHEN** executing `xlex headers report.xlsx "Data" A1:D1`
- **THEN** the aliased command SHALL be executed

#### Scenario: List aliases
- **GIVEN** the command `xlex alias list`
- **WHEN** executed
- **THEN** output SHALL list all defined aliases

### Requirement: Environment Variables

The system SHALL respect environment variables for configuration.

#### Scenario: XLEX_DEFAULT_FORMAT
- **GIVEN** XLEX_DEFAULT_FORMAT=json
- **WHEN** executing `xlex sheet list report.xlsx`
- **THEN** output SHALL be JSON format

#### Scenario: XLEX_NO_COLOR
- **GIVEN** XLEX_NO_COLOR=1
- **WHEN** executing any command
- **THEN** output SHALL not contain color codes

#### Scenario: XLEX_QUIET
- **GIVEN** XLEX_QUIET=1
- **WHEN** executing any command
- **THEN** only errors and essential output SHALL be shown

#### Scenario: XLEX_STRING_CACHE_SIZE
- **GIVEN** XLEX_STRING_CACHE_SIZE=5000
- **WHEN** xlex runs
- **THEN** SharedStrings cache SHALL be limited to 5000 entries

### Requirement: Exit Codes

The system SHALL use consistent exit codes.

#### Scenario: Success
- **GIVEN** a successful operation
- **WHEN** completed
- **THEN** exit code SHALL be 0

#### Scenario: General error
- **GIVEN** an operation that fails
- **WHEN** completed
- **THEN** exit code SHALL be 1

#### Scenario: File not found
- **GIVEN** a command referencing non-existent file
- **WHEN** executed
- **THEN** exit code SHALL be 2

#### Scenario: Invalid arguments
- **GIVEN** a command with invalid arguments
- **WHEN** executed
- **THEN** exit code SHALL be 3

#### Scenario: Permission denied
- **GIVEN** a command on a read-only file
- **WHEN** attempting to write
- **THEN** exit code SHALL be 4

### Requirement: Interactive Mode

The system SHALL support interactive mode via `xlex interactive`.

#### Scenario: Start interactive
- **GIVEN** the command `xlex interactive report.xlsx`
- **WHEN** executed
- **THEN** an interactive REPL SHALL start

#### Scenario: Interactive commands
- **GIVEN** interactive mode is active
- **WHEN** entering `sheet list`
- **THEN** sheets SHALL be listed without specifying file

#### Scenario: Exit interactive
- **GIVEN** interactive mode is active
- **WHEN** entering `exit` or pressing Ctrl+D
- **THEN** interactive mode SHALL exit

#### Scenario: History
- **GIVEN** interactive mode is active
- **WHEN** pressing up arrow
- **THEN** previous commands SHALL be recalled

### Requirement: Man Page Generation

The system SHALL generate man pages via `xlex man`.

#### Scenario: Generate man page
- **GIVEN** the command `xlex man`
- **WHEN** executed
- **THEN** output SHALL be valid man page format

#### Scenario: Generate for subcommand
- **GIVEN** the command `xlex man sheet`
- **WHEN** executed
- **THEN** output SHALL be man page for sheet commands
