#!/bin/bash
# XLEX Installer Script
# Usage: curl -fsSL https://raw.githubusercontent.com/yen0304/xlex/main/install.sh | bash
#
# Environment variables:
#   XLEX_INSTALL_DIR - Installation directory (default: ~/.local/bin)
#   XLEX_VERSION     - Specific version to install (default: latest)

set -e

# Colors
RED='\033[0;31m'
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
BLUE='\033[0;34m'
NC='\033[0m'

REPO="yen0304/xlex"
INSTALL_DIR="${XLEX_INSTALL_DIR:-$HOME/.local/bin}"
VERSION="${XLEX_VERSION:-latest}"

info() {
    echo -e "${BLUE}[INFO]${NC} $1"
}

warn() {
    echo -e "${YELLOW}[WARN]${NC} $1"
}

error() {
    echo -e "${RED}[ERROR]${NC} $1"
    exit 1
}

success() {
    echo -e "${GREEN}[OK]${NC} $1"
}

detect_os() {
    case "$(uname -s)" in
        Linux*)     OS="linux";;
        Darwin*)    OS="macos";;
        CYGWIN*|MINGW*|MSYS*) OS="windows";;
        *)          error "Unsupported operating system: $(uname -s)";;
    esac
}

detect_arch() {
    case "$(uname -m)" in
        x86_64|amd64)   ARCH="x86_64";;
        aarch64|arm64)  ARCH="aarch64";;
        *)              error "Unsupported architecture: $(uname -m)";;
    esac
}

get_latest_version() {
    if [ "$VERSION" = "latest" ]; then
        VERSION=$(curl -sL "https://api.github.com/repos/${REPO}/releases/latest" | grep '"tag_name":' | sed -E 's/.*"([^"]+)".*/\1/')
        if [ -z "$VERSION" ]; then
            error "Failed to get latest version"
        fi
    fi
}

get_download_url() {
    local filename
    case "$OS" in
        linux)
            filename="xlex-linux-${ARCH}.tar.gz"
            ;;
        macos)
            filename="xlex-macos-${ARCH}.tar.gz"
            ;;
        windows)
            filename="xlex-windows-${ARCH}.zip"
            ;;
    esac
    DOWNLOAD_URL="https://github.com/${REPO}/releases/download/${VERSION}/${filename}"
}

download_and_install() {
    local tmp_dir
    tmp_dir=$(mktemp -d)
    trap "rm -rf $tmp_dir" EXIT

    info "Downloading xlex ${VERSION} for ${OS}/${ARCH}..."
    
    if command -v curl &> /dev/null; then
        curl -fsSL "$DOWNLOAD_URL" -o "$tmp_dir/xlex.tar.gz"
    elif command -v wget &> /dev/null; then
        wget -qO "$tmp_dir/xlex.tar.gz" "$DOWNLOAD_URL"
    else
        error "Neither curl nor wget found. Please install one of them."
    fi

    info "Extracting..."
    if [ "$OS" = "windows" ]; then
        unzip -q "$tmp_dir/xlex.tar.gz" -d "$tmp_dir"
    else
        tar -xzf "$tmp_dir/xlex.tar.gz" -C "$tmp_dir"
    fi

    info "Installing to ${INSTALL_DIR}..."
    mkdir -p "$INSTALL_DIR"
    
    if [ "$OS" = "windows" ]; then
        mv "$tmp_dir/xlex.exe" "$INSTALL_DIR/"
    else
        mv "$tmp_dir/xlex" "$INSTALL_DIR/"
        chmod +x "$INSTALL_DIR/xlex"
    fi

    success "xlex ${VERSION} installed successfully!"
}

verify_installation() {
    if [ -x "$INSTALL_DIR/xlex" ]; then
        local version_output
        version_output=$("$INSTALL_DIR/xlex" --version 2>/dev/null || echo "unknown")
        success "Installed: $version_output"
    else
        warn "Installation completed but binary verification failed"
    fi
}

check_path() {
    if [[ ":$PATH:" != *":$INSTALL_DIR:"* ]]; then
        warn "Add ${INSTALL_DIR} to your PATH:"
        echo ""
        echo "  # For bash (add to ~/.bashrc):"
        echo "  export PATH=\"\$PATH:${INSTALL_DIR}\""
        echo ""
        echo "  # For zsh (add to ~/.zshrc):"
        echo "  export PATH=\"\$PATH:${INSTALL_DIR}\""
        echo ""
        echo "  # For fish (add to ~/.config/fish/config.fish):"
        echo "  set -gx PATH \$PATH ${INSTALL_DIR}"
        echo ""
    fi
}

main() {
    echo ""
    echo "╔═══════════════════════════════════════╗"
    echo "║         XLEX Installer                ║"
    echo "║   Excel CLI Tool for Developers       ║"
    echo "╚═══════════════════════════════════════╝"
    echo ""

    detect_os
    detect_arch
    get_latest_version
    get_download_url

    info "OS: $OS, Arch: $ARCH"
    info "Version: $VERSION"
    info "Install directory: $INSTALL_DIR"
    echo ""

    download_and_install
    verify_installation
    check_path

    echo ""
    success "Installation complete! Run 'xlex --help' to get started."
    echo ""
}

main "$@"
