# Installation

XLEX can be installed on Linux, macOS, and Windows using several methods.

## Quick Install

The fastest way to install XLEX:

=== "Linux/macOS"

    ```bash
    curl -fsSL https://raw.githubusercontent.com/yen0304/xlex/main/install.sh | bash
    ```

=== "Windows (PowerShell)"

    ```powershell
    irm https://raw.githubusercontent.com/yen0304/xlex/main/install.ps1 | iex
    ```

## Package Managers

### Cargo (Rust)

If you have Rust installed:

```bash
cargo install xlex-cli
```

### Homebrew (macOS/Linux)

```bash
brew install user/tap/xlex
```

### npm/npx

Using npm (requires Node.js 14+):

```bash
# Install globally
npm install -g xlex-cli

# Or use with npx (no install)
npx xlex-cli --help
```

## Manual Installation

### Download Binary

1. Go to the [Releases page](https://github.com/yen0304/xlex/releases)
2. Download the appropriate archive for your platform:
   - `xlex-linux-x86_64.tar.gz` - Linux (Intel/AMD)
   - `xlex-linux-aarch64.tar.gz` - Linux (ARM64)
   - `xlex-macos-x86_64.tar.gz` - macOS (Intel)
   - `xlex-macos-aarch64.tar.gz` - macOS (Apple Silicon)
   - `xlex-windows-x86_64.zip` - Windows

3. Extract and add to your PATH:

=== "Linux/macOS"

    ```bash
    tar -xzf xlex-linux-x86_64.tar.gz
    sudo mv xlex /usr/local/bin/
    ```

=== "Windows"

    Extract `xlex.exe` and add its location to your PATH environment variable.

### Build from Source

Requirements:
- Rust 1.70+
- Git

```bash
git clone https://github.com/yen0304/xlex.git
cd xlex
cargo build --release
# Binary is at target/release/xlex
```

## Verify Installation

```bash
xlex --version
# xlex 0.1.0

xlex --help
# Shows available commands
```

## Shell Completions

Generate shell completions for better tab-completion:

=== "Bash"

    ```bash
    xlex completion bash > ~/.local/share/bash-completion/completions/xlex
    ```

=== "Zsh"

    ```bash
    xlex completion zsh > ~/.zfunc/_xlex
    # Add to .zshrc: fpath+=~/.zfunc
    ```

=== "Fish"

    ```bash
    xlex completion fish > ~/.config/fish/completions/xlex.fish
    ```

=== "PowerShell"

    ```powershell
    xlex completion powershell > $HOME\Documents\PowerShell\Modules\xlex.ps1
    ```

## Updating

To update XLEX to the latest version:

=== "Cargo"

    ```bash
    cargo install xlex-cli --force
    ```

=== "Homebrew"

    ```bash
    brew upgrade xlex
    ```

=== "npm"

    ```bash
    npm update -g xlex-cli
    ```

## Uninstalling

=== "Cargo"

    ```bash
    cargo uninstall xlex-cli
    ```

=== "Homebrew"

    ```bash
    brew uninstall xlex
    ```

=== "npm"

    ```bash
    npm uninstall -g xlex-cli
    ```

=== "Manual"

    ```bash
    rm /usr/local/bin/xlex
    ```
