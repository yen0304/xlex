# Homebrew formula for XLEX
# To use: brew tap user/tap && brew install xlex
# Or: brew install user/tap/xlex

class Xlex < Formula
  desc "CLI-first streaming Excel manipulation tool for developers"
  homepage "https://github.com/yen0304/xlex"
  version "0.1.0"
  license "MIT"

  on_macos do
    on_arm do
      url "https://github.com/yen0304/xlex/releases/download/v#{version}/xlex-macos-aarch64.tar.gz"
      sha256 "PLACEHOLDER_SHA256_MACOS_ARM64"
    end
    on_intel do
      url "https://github.com/yen0304/xlex/releases/download/v#{version}/xlex-macos-x86_64.tar.gz"
      sha256 "PLACEHOLDER_SHA256_MACOS_X86_64"
    end
  end

  on_linux do
    on_arm do
      url "https://github.com/yen0304/xlex/releases/download/v#{version}/xlex-linux-aarch64.tar.gz"
      sha256 "PLACEHOLDER_SHA256_LINUX_ARM64"
    end
    on_intel do
      url "https://github.com/yen0304/xlex/releases/download/v#{version}/xlex-linux-x86_64.tar.gz"
      sha256 "PLACEHOLDER_SHA256_LINUX_X86_64"
    end
  end

  def install
    bin.install "xlex"

    # Generate shell completions
    generate_completions_from_executable(bin/"xlex", "completion")
  end

  test do
    # Test version command
    assert_match version.to_s, shell_output("#{bin}/xlex --version")

    # Test create and info
    system bin/"xlex", "create", "test.xlsx"
    assert_predicate testpath/"test.xlsx", :exist?

    output = shell_output("#{bin}/xlex info test.xlsx")
    assert_match "Sheets:", output
  end
end
