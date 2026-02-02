# Tasks: XLEX CLI Excel Engine

## Phase 1: Project Foundation

- [x] 1.1 Initialize Cargo workspace with root Cargo.toml
- [x] 1.2 Create crates/xlex-core/ with Cargo.toml
- [x] 1.3 Create crates/xlex-cli/ with Cargo.toml
- [x] 1.4 Configure dependencies (zip, quick-xml, clap, serde)
- [x] 1.5 Set up CI/CD pipeline (GitHub Actions)
- [x] 1.6 Configure rustfmt and clippy
- [x] 1.7 Create test fixture directory with sample xlsx files
- [x] 1.8 Set up benchmark infrastructure
- [x] 1.9 Create initial GitHub repository structure

## Phase 2: Core Engine (xlex-core)

- [x] 2.1 Implement ZIP archive reader with streaming support
- [x] 2.2 Implement XML SAX parser wrapper for xlsx elements
- [x] 2.3 Implement SharedStrings lazy loader with LRU cache
- [x] 2.4 Implement Style registry parser
- [x] 2.5 Implement Workbook structure parser (workbook.xml)
- [x] 2.6 Implement Sheet XML streaming parser
- [x] 2.7 Implement Cell reference parser (A1 notation)
- [x] 2.8 Implement Range parser (A1:B10 notation)
- [x] 2.9 Implement copy-on-write ZIP writer
- [x] 2.10 Define public API for xlex-core
- [x] 2.11 Write unit tests for core engine components

## Phase 3: Data Models (xlex-core)

- [x] 3.1 Define Workbook struct and traits
- [x] 3.2 Define Sheet struct with streaming iterator
- [x] 3.3 Define Cell struct with type variants
- [x] 3.4 Define CellValue enum (String, Number, Boolean, Formula, Error)
- [x] 3.5 Define Style struct and StyleRegistry
- [x] 3.6 Define Range struct with iteration support
- [x] 3.7 Define Error types with codes
- [x] 3.8 Write unit tests for data models

## Phase 4: Workbook Operations

- [x] 4.1 Implement `xlex info` command
- [x] 4.2 Implement `xlex validate` command
- [x] 4.3 Implement `xlex clone` command
- [x] 4.4 Implement `xlex create` command
- [x] 4.5 Implement `xlex props get` command
- [x] 4.6 Implement `xlex props set` command
- [x] 4.7 Implement `xlex stats` command
- [x] 4.8 Write integration tests for workbook operations

## Phase 5: Sheet Operations

- [x] 5.1 Implement `xlex sheet list` command
- [x] 5.2 Implement `xlex sheet add` command
- [x] 5.3 Implement `xlex sheet remove` command
- [x] 5.4 Implement `xlex sheet rename` command
- [x] 5.5 Implement `xlex sheet copy` command
- [x] 5.6 Implement `xlex sheet move` command
- [x] 5.7 Implement `xlex sheet hide` command
- [x] 5.8 Implement `xlex sheet unhide` command
- [x] 5.9 Implement `xlex sheet info` command
- [x] 5.10 Implement `xlex sheet active` command
- [x] 5.11 Write integration tests for sheet operations

## Phase 6: Cell Operations

- [x] 6.1 Implement `xlex cell get` command
- [x] 6.2 Implement `xlex cell set` command
- [x] 6.3 Implement `xlex cell formula` command
- [x] 6.4 Implement `xlex cell clear` command
- [x] 6.5 Implement `xlex cell type` command
- [x] 6.6 Implement `xlex cell batch` command (stdin)
- [x] 6.7 Implement `xlex cell comment` subcommands (get/set/remove/list)
- [x] 6.8 Implement `xlex cell link` subcommands (get/set/remove)
- [x] 6.9 Write integration tests for cell operations

## Phase 7: Row Operations

- [x] 7.1 Implement `xlex row get` command
- [x] 7.2 Implement `xlex row append` command
- [x] 7.3 Implement `xlex row insert` command
- [x] 7.4 Implement `xlex row delete` command
- [x] 7.5 Implement `xlex row copy` command
- [x] 7.6 Implement `xlex row move` command
- [x] 7.7 Implement `xlex row height` command
- [x] 7.8 Implement `xlex row hide` command
- [x] 7.9 Implement `xlex row unhide` command
- [x] 7.10 Implement `xlex row find` command
- [x] 7.11 Write integration tests for row operations

## Phase 8: Column Operations

- [x] 8.1 Implement `xlex column get` command
- [x] 8.2 Implement `xlex column insert` command
- [x] 8.3 Implement `xlex column delete` command
- [x] 8.4 Implement `xlex column copy` command
- [x] 8.5 Implement `xlex column move` command
- [x] 8.6 Implement `xlex column width` command
- [x] 8.7 Implement `xlex column hide` command
- [x] 8.8 Implement `xlex column unhide` command
- [x] 8.9 Implement `xlex column header` command
- [x] 8.10 Implement `xlex column find` command
- [x] 8.11 Implement `xlex column stats` command
- [x] 8.12 Write integration tests for column operations

## Phase 9: Range Operations

- [x] 9.1 Implement `xlex range get` command
- [x] 9.2 Implement `xlex range copy` command
- [x] 9.3 Implement `xlex range move` command
- [x] 9.4 Implement `xlex range clear` command
- [x] 9.5 Implement `xlex range fill` command
- [x] 9.6 Implement `xlex range merge` command
- [x] 9.7 Implement `xlex range unmerge` command
- [x] 9.8 Implement `xlex range name` command (named ranges)
- [x] 9.9 Implement `xlex range names` command (list named ranges)
- [x] 9.10 Implement `xlex range validate` command
- [x] 9.11 Implement `xlex range sort` command
- [x] 9.12 Implement `xlex range filter` command
- [x] 9.13 Write integration tests for range operations

## Phase 10: Style Operations

- [x] 10.1 Implement `xlex style list` command
- [x] 10.2 Implement `xlex style get` command
- [x] 10.3 Implement `xlex range style` command with all flags
- [x] 10.4 Implement `xlex range border` command
- [x] 10.5 Implement `xlex style preset` subcommands (list/apply/create/delete)
- [x] 10.6 Implement `xlex style copy` command
- [x] 10.7 Implement `xlex style clear` command
- [x] 10.8 Implement `xlex style condition` command
- [x] 10.9 Implement `xlex style freeze` command
- [x] 10.10 Write integration tests for style operations

## Phase 11: Import/Export

- [x] 11.1 Implement `xlex to csv` command
- [x] 11.2 Implement `xlex to json` command
- [x] 11.3 Implement `xlex to ndjson` command
- [x] 11.4 Implement `xlex to meta` command
- [x] 11.5 Implement `xlex from csv` command
- [x] 11.6 Implement `xlex from json` command
- [x] 11.7 Implement `xlex from ndjson` command
- [x] 11.8 Implement `xlex convert` command
- [x] 11.9 Implement streaming for large exports
- [x] 11.10 Implement multi-sheet export (--all flag)
- [x] 11.11 Write integration tests for import/export

## Phase 12: Formula Operations

- [x] 12.1 Implement formula syntax parser/validator
- [x] 12.2 Implement `xlex formula validate` command
- [x] 12.3 Implement `xlex formula list` command
- [x] 12.4 Implement `xlex formula stats` command
- [x] 12.5 Implement `xlex formula refs` command (find dependents/precedents)
- [x] 12.6 Implement `xlex formula replace` command
- [x] 12.7 Implement circular reference detection
- [x] 12.8 Write integration tests for formula operations

## Phase 13: Template Operations

- [x] 13.1 Implement template placeholder parser (Handlebars-like syntax)
- [x] 13.2 Implement `xlex template apply` command
- [x] 13.3 Implement `xlex template validate` command
- [x] 13.4 Implement `xlex template init` command
- [x] 13.5 Implement `xlex template preview` command
- [x] 13.6 Implement template loops ({{#each}})
- [x] 13.7 Implement template conditionals ({{#if}})
- [x] 13.8 Implement template filters (date, number formatting)
- [x] 13.9 Implement batch template processing (--per-record)
- [x] 13.10 Implement special markers (row-repeat, sheet-repeat, image)
- [x] 13.11 Write integration tests for template operations

## Phase 14: CLI Polish

- [x] 14.1 Implement global flags (--quiet, --verbose, --format, --no-color)
- [x] 14.2 Implement `xlex completion` command (bash/zsh/fish/powershell)
- [x] 14.3 Implement `xlex help` with examples
- [x] 14.4 Implement `xlex version` command
- [x] 14.5 Implement `xlex config` subcommands (show/get/set/reset/init/validate)
- [x] 14.6 Implement project config file loading (.xlex.yml)
- [x] 14.7 Implement config precedence (CLI > env > project > user > default)
- [x] 14.8 Implement `xlex batch` command
- [x] 14.9 Implement `xlex alias` subcommands (add/list/remove)
- [x] 14.10 Implement `xlex interactive` mode
- [x] 14.11 Add colored output support
- [x] 14.12 Implement progress indicators for long operations
- [x] 14.13 Implement `xlex man` command

## Phase 15: Error Handling

- [x] 15.1 Define all error codes (XLEX_E001 - XLEX_E099)
- [x] 15.2 Implement error formatting (human + machine readable)
- [x] 15.3 Implement --json-errors flag
- [x] 15.4 Add error recovery suggestions
- [x] 15.5 Implement error logging (XLEX_LOG_FILE)
- [x] 15.6 Write error handling tests

## Phase 16: Performance Validation

- [x] 16.1 Create 200MB test fixture
- [x] 16.2 Benchmark sheet list (<100ms target)
- [x] 16.3 Benchmark column read (<300ms target)
- [x] 16.4 Benchmark cell update (<200ms target)
- [x] 16.5 Benchmark 10k row append (<1s target)
- [x] 16.6 Profile memory usage
- [x] 16.7 Optimize hot paths if needed

## Phase 17: GitHub & Community

- [x] 17.1 Create README.md with badges, features, quick start
- [x] 17.2 Create LICENSE (MIT)
- [x] 17.3 Create CODE_OF_CONDUCT.md (Contributor Covenant)
- [x] 17.4 Create CONTRIBUTING.md with development setup
- [x] 17.5 Create SECURITY.md with vulnerability reporting
- [x] 17.6 Create CHANGELOG.md (Keep a Changelog format)
- [x] 17.7 Create .github/ISSUE_TEMPLATE/bug_report.yml
- [x] 17.8 Create .github/ISSUE_TEMPLATE/feature_request.yml
- [x] 17.9 Create .github/ISSUE_TEMPLATE/question.yml
- [x] 17.10 Create .github/ISSUE_TEMPLATE/config.yml
- [x] 17.11 Create .github/PULL_REQUEST_TEMPLATE.md
- [x] 17.12 Create .github/FUNDING.yml
- [x] 17.13 Create .github/dependabot.yml

## Phase 18: CI/CD Workflows

- [x] 18.1 Create .github/workflows/ci.yml (test, lint, build)
- [x] 18.2 Configure matrix testing (Linux, macOS, Windows)
- [x] 18.3 Create .github/workflows/release.yml
- [x] 18.4 Configure release artifact builds for all platforms
- [x] 18.5 Configure automatic changelog generation
- [x] 18.6 Set up code coverage reporting

## Phase 19: Distribution

- [x] 19.1 Configure GitHub Releases workflow
- [x] 19.2 Create install.sh script for curl installation
- [x] 19.3 Create npm wrapper package for npx
- [x] 19.4 Create Homebrew formula (xlex.rb)
- [x] ~~19.5 Publish to crates.io~~ (deferred - requires actual publish)
- [x] ~~19.6 Set up xlex.sh domain~~ (deferred - requires domain setup)
- [x] 19.7 Create SHA256 checksums for releases

## Phase 20: Documentation - Getting Started

- [x] 20.1 Create docs/ directory structure
- [x] 20.2 Create mkdocs.yml configuration
- [x] 20.3 Write docs/index.md (home page)
- [x] 20.4 Write docs/getting-started/installation.md
- [x] 20.5 Write docs/getting-started/quick-start.md
- [x] ~~20.6 Write docs/getting-started/first-steps.md~~ (removed - lean docs)
- [x] ~~20.7 Create sample.xlsx for tutorials~~ (removed - lean docs)

## Phase 21: Documentation - Reference (Lean)

- [x] 21.1 Write docs/reference/cli-reference.md (complete CLI synopsis)
- [x] 21.2 Write docs/reference/exit-codes.md
- [x] ~~21.3-21.12 Individual command docs~~ (removed - use --help and cli-reference.md)

## Phase 22-25: REMOVED (Lean Documentation Strategy)

The following phases were removed to keep documentation minimal and maintainable:

- ~~Phase 22: Guides~~ - Users can rely on --help and examples in README
- ~~Phase 23: Cookbook~~ - Examples included in quick-start.md
- ~~Phase 24: Development docs~~ - Use CONTRIBUTING.md instead
- ~~Phase 25: Detailed reference~~ - CLI --help is comprehensive

**Final documentation structure (5 files):**
```
docs/
├── index.md                    # Home page
├── getting-started/
│   ├── installation.md         # All install methods
│   └── quick-start.md          # Quick examples
└── reference/
    ├── cli-reference.md        # Full command reference
    └── exit-codes.md           # Exit codes
```

## Phase 26: Documentation Deployment

- [x] 26.1 Configure GitHub Pages deployment
- [x] ~~26.2 Set up docs.xlex.sh domain~~ (optional - use GitHub Pages URL)
- [x] 26.3 Configure automatic docs build on release
- [x] ~~26.4 Add documentation search~~ (not needed for 5 pages)

## Dependencies

- Phase 2 (Core Engine) blocks all operation phases (4-13)
- Phase 3 (Data Models) blocks all operation phases (4-13)
- Phase 4-13 can be parallelized after Phase 2-3 complete
- Phase 14-15 can start after Phase 4 complete
- Phase 16 requires all operations complete
- Phase 17-18 can start at project initialization
- Phase 19 requires Phase 16 and 18 complete
- Phase 20-21 (Documentation) minimal set after features complete
- Phase 26 (Deployment) after Phase 20-21 complete

## Parallelization Opportunities

The following can be done in parallel:
- Phase 17-18 (GitHub/CI) with Phase 2-3 (Core Engine)
- Phase 4-13 (Operations) with each other after Phase 2-3
- Phase 12 (Formula) and Phase 13 (Template) can be done in parallel

## Documentation Philosophy (Lean Approach)

For a new open source project, minimal documentation is preferred:
- Comprehensive `--help` in CLI is the primary documentation
- README.md covers quick start and common examples
- Only 5 doc pages to reduce maintenance burden
- Documentation grows with community needs, not upfront
