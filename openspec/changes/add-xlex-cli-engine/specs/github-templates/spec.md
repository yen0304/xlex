# GitHub Templates & Community Files Specification

## ADDED Requirements

### Requirement: README.md

The project SHALL provide a comprehensive README.md at the repository root.

#### Scenario: README structure
- **GIVEN** a user visits the GitHub repository
- **WHEN** viewing README.md
- **THEN** the README SHALL include:
  - Project logo/banner
  - One-line description
  - Badges (CI status, version, license, downloads)
  - Feature highlights
  - Quick start installation
  - Basic usage examples
  - Link to full documentation
  - Contributing link
  - License information

#### Scenario: Installation section
- **GIVEN** a user reads the README
- **WHEN** looking for installation instructions
- **THEN** the README SHALL show all installation methods:
  - curl | sh (recommended)
  - Homebrew
  - cargo install
  - npx
  - GitHub Releases download

#### Scenario: Quick examples
- **GIVEN** a user reads the README
- **WHEN** looking for usage examples
- **THEN** the README SHALL include at least 5 common use cases:
  - List sheets
  - Read cell value
  - Update cell
  - Export to CSV
  - Import from CSV

### Requirement: CONTRIBUTING.md

The project SHALL provide contribution guidelines in CONTRIBUTING.md.

#### Scenario: Contribution guide structure
- **GIVEN** a contributor wants to help
- **WHEN** reading CONTRIBUTING.md
- **THEN** the guide SHALL include:
  - Code of Conduct reference
  - How to report bugs
  - How to suggest features
  - Development setup instructions
  - Code style guidelines
  - Pull request process
  - Commit message conventions

#### Scenario: Development setup
- **GIVEN** a developer wants to contribute
- **WHEN** following setup instructions
- **THEN** the guide SHALL explain:
  - Required Rust version
  - How to clone and build
  - How to run tests
  - How to run benchmarks
  - IDE recommendations

#### Scenario: Pull request checklist
- **GIVEN** a contributor submits a PR
- **WHEN** reviewing the PR process
- **THEN** the guide SHALL specify:
  - Tests required
  - Documentation updates required
  - Changelog entry required
  - Code review process
  - CI checks that must pass

### Requirement: CODE_OF_CONDUCT.md

The project SHALL provide a Code of Conduct.

#### Scenario: Code of Conduct content
- **GIVEN** a community member
- **WHEN** reading CODE_OF_CONDUCT.md
- **THEN** it SHALL include:
  - Expected behavior standards
  - Unacceptable behavior examples
  - Enforcement responsibilities
  - Reporting instructions
  - Contact information

#### Scenario: Contributor Covenant
- **GIVEN** the project adopts a standard CoC
- **WHEN** CODE_OF_CONDUCT.md is created
- **THEN** it SHALL be based on Contributor Covenant v2.1

### Requirement: LICENSE

The project SHALL include a LICENSE file.

#### Scenario: License file
- **GIVEN** a user checks the license
- **WHEN** viewing LICENSE
- **THEN** it SHALL contain the full MIT license text
- **AND** include the correct copyright year and holder

#### Scenario: License in README
- **GIVEN** a user reads the README
- **WHEN** looking for license info
- **THEN** the README SHALL state "MIT License" with link to LICENSE file

### Requirement: CHANGELOG.md

The project SHALL maintain a changelog.

#### Scenario: Changelog format
- **GIVEN** a user checks release history
- **WHEN** reading CHANGELOG.md
- **THEN** it SHALL follow Keep a Changelog format:
  - Versions in reverse chronological order
  - Categories: Added, Changed, Deprecated, Removed, Fixed, Security
  - Links to version diffs

#### Scenario: Unreleased section
- **GIVEN** development is ongoing
- **WHEN** changes are merged
- **THEN** an [Unreleased] section SHALL track pending changes

### Requirement: SECURITY.md

The project SHALL provide security policy.

#### Scenario: Security policy content
- **GIVEN** a security researcher finds a vulnerability
- **WHEN** reading SECURITY.md
- **THEN** it SHALL include:
  - Supported versions table
  - How to report vulnerabilities
  - Expected response timeline
  - Disclosure policy

#### Scenario: Responsible disclosure
- **GIVEN** a vulnerability is reported
- **WHEN** following the security policy
- **THEN** the policy SHALL specify private reporting via GitHub Security Advisories

### Requirement: Bug Report Issue Template

The project SHALL provide a bug report template at .github/ISSUE_TEMPLATE/bug_report.yml.

#### Scenario: Bug report fields
- **GIVEN** a user reports a bug
- **WHEN** creating a new issue
- **THEN** the template SHALL request:
  - Bug description
  - Steps to reproduce
  - Expected behavior
  - Actual behavior
  - XLEX version
  - Operating system
  - xlsx file sample (if applicable)
  - Error message/output

#### Scenario: Bug report validation
- **GIVEN** a user fills the bug report
- **WHEN** submitting
- **THEN** required fields SHALL be enforced:
  - Description (required)
  - Steps to reproduce (required)
  - Version (required)
  - OS (required)

### Requirement: Feature Request Issue Template

The project SHALL provide a feature request template at .github/ISSUE_TEMPLATE/feature_request.yml.

#### Scenario: Feature request fields
- **GIVEN** a user suggests a feature
- **WHEN** creating a new issue
- **THEN** the template SHALL request:
  - Feature summary
  - Problem it solves
  - Proposed solution
  - Alternative solutions considered
  - Additional context

#### Scenario: Feature request labels
- **GIVEN** a feature request is submitted
- **WHEN** the issue is created
- **THEN** it SHALL be automatically labeled "enhancement"

### Requirement: Question/Support Issue Template

The project SHALL provide a question template at .github/ISSUE_TEMPLATE/question.yml.

#### Scenario: Question fields
- **GIVEN** a user has a question
- **WHEN** creating a new issue
- **THEN** the template SHALL request:
  - Question summary
  - What you've tried
  - Relevant documentation checked
  - XLEX version

#### Scenario: Question labels
- **GIVEN** a question is submitted
- **WHEN** the issue is created
- **THEN** it SHALL be automatically labeled "question"

### Requirement: Issue Template Config

The project SHALL provide issue template configuration at .github/ISSUE_TEMPLATE/config.yml.

#### Scenario: Template chooser
- **GIVEN** a user clicks "New Issue"
- **WHEN** the template chooser appears
- **THEN** it SHALL show:
  - Bug Report
  - Feature Request
  - Question
  - Link to Discussions (if enabled)

#### Scenario: Blank issues
- **GIVEN** the template config
- **WHEN** configured
- **THEN** blank issues SHALL be disabled to encourage template use

### Requirement: Pull Request Template

The project SHALL provide a PR template at .github/PULL_REQUEST_TEMPLATE.md.

#### Scenario: PR template content
- **GIVEN** a contributor opens a PR
- **WHEN** the PR form appears
- **THEN** the template SHALL include:
  - Description of changes
  - Related issue reference
  - Type of change (bug fix, feature, docs, etc.)
  - Checklist:
    - [ ] Tests added/updated
    - [ ] Documentation updated
    - [ ] CHANGELOG.md updated
    - [ ] `cargo fmt` run
    - [ ] `cargo clippy` passes

### Requirement: GitHub Actions CI Workflow

The project SHALL provide CI workflow at .github/workflows/ci.yml.

#### Scenario: CI triggers
- **GIVEN** the CI workflow
- **WHEN** configured
- **THEN** it SHALL trigger on:
  - Push to main
  - Pull requests to main
  - Manual dispatch

#### Scenario: CI jobs
- **GIVEN** CI runs
- **WHEN** executing
- **THEN** it SHALL run:
  - cargo fmt --check
  - cargo clippy -- -D warnings
  - cargo test
  - cargo build --release
  - Integration tests

#### Scenario: Matrix testing
- **GIVEN** CI runs
- **WHEN** testing
- **THEN** it SHALL test on:
  - Ubuntu latest
  - macOS latest
  - Windows latest

### Requirement: Release Workflow

The project SHALL provide release workflow at .github/workflows/release.yml.

#### Scenario: Release triggers
- **GIVEN** the release workflow
- **WHEN** a tag is pushed (v*)
- **THEN** it SHALL:
  - Build release binaries for all platforms
  - Create GitHub Release
  - Upload binaries as assets
  - Generate release notes from CHANGELOG

#### Scenario: Release artifacts
- **GIVEN** a release is created
- **WHEN** binaries are built
- **THEN** artifacts SHALL include:
  - xlex-linux-x86_64.tar.gz
  - xlex-linux-aarch64.tar.gz
  - xlex-darwin-x86_64.tar.gz
  - xlex-darwin-aarch64.tar.gz
  - xlex-windows-x86_64.zip
  - SHA256 checksums file

### Requirement: Dependabot Configuration

The project SHALL provide Dependabot config at .github/dependabot.yml.

#### Scenario: Dependabot settings
- **GIVEN** Dependabot is configured
- **WHEN** checking for updates
- **THEN** it SHALL:
  - Check cargo dependencies weekly
  - Check GitHub Actions weekly
  - Group minor/patch updates
  - Auto-label PRs

### Requirement: Funding Configuration

The project SHALL provide funding config at .github/FUNDING.yml.

#### Scenario: Funding options
- **GIVEN** a user wants to sponsor
- **WHEN** clicking "Sponsor"
- **THEN** available options SHALL be displayed:
  - GitHub Sponsors (if configured)
  - Open Collective (if configured)
  - Custom links (if configured)

### Requirement: Repository Settings Documentation

The project SHALL document recommended repository settings.

#### Scenario: Branch protection
- **GIVEN** repository settings documentation
- **WHEN** configuring the repo
- **THEN** it SHALL recommend:
  - Require PR reviews before merging
  - Require status checks to pass
  - Require linear history
  - Do not allow force pushes to main

#### Scenario: Repository features
- **GIVEN** repository settings documentation
- **WHEN** configuring the repo
- **THEN** it SHALL recommend enabling:
  - Issues
  - Discussions
  - Projects (optional)
  - Wiki (disabled, use docs instead)

### Requirement: Community Health Files

The project SHALL achieve GitHub Community Standards compliance.

#### Scenario: Community profile
- **GIVEN** the repository
- **WHEN** checking Community Standards
- **THEN** all items SHALL be green:
  - Description
  - README
  - Code of conduct
  - Contributing
  - License
  - Security policy
  - Issue templates
  - Pull request template
