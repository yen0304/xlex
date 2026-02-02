#!/bin/bash
# XLEX CLI Integration Test Script
# This script tests all implemented CLI commands

# Don't exit on error - we want to continue testing
# set -e

# Colors for output
RED='\033[0;31m'
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
BLUE='\033[0;34m'
NC='\033[0m' # No Color

# Test directory
TEST_DIR="/tmp/xlex_test_$$"
XLEX="${XLEX_BIN:-./target/release/xlex}"

# Counters
PASSED=0
FAILED=0
SKIPPED=0

# Create test directory
mkdir -p "$TEST_DIR"

# Helper functions
log_info() {
    echo -e "${BLUE}[INFO]${NC} $1"
}

log_pass() {
    echo -e "${GREEN}[PASS]${NC} $1"
    ((PASSED++))
}

log_fail() {
    echo -e "${RED}[FAIL]${NC} $1"
    ((FAILED++))
}

log_skip() {
    echo -e "${YELLOW}[SKIP]${NC} $1"
    ((SKIPPED++))
}

# Run a test
run_test() {
    local name="$1"
    local cmd="$2"
    local expected="$3"
    
    log_info "Testing: $name"
    
    if output=$(eval "$cmd" 2>&1); then
        if [[ -z "$expected" ]] || [[ "$output" == *"$expected"* ]]; then
            log_pass "$name"
            return 0
        else
            log_fail "$name - Expected: '$expected', Got: '$output'"
            return 1
        fi
    else
        log_fail "$name - Command failed: $output"
        return 1
    fi
}

# Run a test expecting failure
run_test_fail() {
    local name="$1"
    local cmd="$2"
    
    log_info "Testing (expect fail): $name"
    
    if output=$(eval "$cmd" 2>&1); then
        log_fail "$name - Expected failure but succeeded"
        return 1
    else
        log_pass "$name"
        return 0
    fi
}

cleanup() {
    log_info "Cleaning up test directory..."
    rm -rf "$TEST_DIR"
}

trap cleanup EXIT

echo "========================================"
echo "       XLEX CLI Integration Tests"
echo "========================================"
echo ""

# Check if binary exists
if [[ ! -f "$XLEX" ]]; then
    echo -e "${RED}Error: xlex binary not found at $XLEX${NC}"
    echo "Please run: cargo build --release"
    exit 1
fi

log_info "Using xlex binary: $XLEX"
log_info "Test directory: $TEST_DIR"
echo ""

# ===== Phase 1: Basic Commands =====
echo "--- Phase 1: Basic Commands ---"

run_test "version" "$XLEX version" "xlex"
run_test "help" "$XLEX --help" "Usage:"

# ===== Phase 2: Workbook Operations =====
echo ""
echo "--- Phase 2: Workbook Operations ---"

run_test "create workbook" "$XLEX create $TEST_DIR/test.xlsx" "Created"
run_test "info" "$XLEX info $TEST_DIR/test.xlsx" "Sheet1"
run_test "validate" "$XLEX validate $TEST_DIR/test.xlsx" "valid"
run_test "clone" "$XLEX clone $TEST_DIR/test.xlsx $TEST_DIR/test_copy.xlsx" "Cloned"
run_test "stats" "$XLEX stats $TEST_DIR/test.xlsx" "Sheets:"
run_test "props set" "$XLEX props set $TEST_DIR/test.xlsx title 'Test Workbook'" "Set title"
run_test "props get" "$XLEX props get $TEST_DIR/test.xlsx" "Test Workbook"

# ===== Phase 3: Sheet Operations =====
echo ""
echo "--- Phase 3: Sheet Operations ---"

run_test "sheet list" "$XLEX sheet list $TEST_DIR/test.xlsx" "Sheet1"
run_test "sheet add" "$XLEX sheet add $TEST_DIR/test.xlsx Sheet2" "Added sheet"
run_test "sheet add 2" "$XLEX sheet add $TEST_DIR/test.xlsx Sheet3" "Added sheet"
run_test "sheet rename" "$XLEX sheet rename $TEST_DIR/test.xlsx Sheet2 MySheet" "Renamed"
run_test "sheet copy" "$XLEX sheet copy $TEST_DIR/test.xlsx Sheet1 Sheet1_Copy" "Copied"
run_test "sheet move" "$XLEX sheet move $TEST_DIR/test.xlsx Sheet3 0" "Moved"
run_test "sheet hide" "$XLEX sheet hide $TEST_DIR/test.xlsx MySheet" "Hid sheet"
run_test "sheet unhide" "$XLEX sheet unhide $TEST_DIR/test.xlsx MySheet" "Unhid sheet"
run_test "sheet info" "$XLEX sheet info $TEST_DIR/test.xlsx Sheet1" "Name: Sheet1"
run_test "sheet active" "$XLEX sheet active $TEST_DIR/test.xlsx MySheet" "Set active"
run_test "sheet remove" "$XLEX sheet remove $TEST_DIR/test.xlsx Sheet1_Copy" "Removed"

# ===== Phase 4: Cell Operations =====
echo ""
echo "--- Phase 4: Cell Operations ---"

run_test "cell set string" "$XLEX cell set $TEST_DIR/test.xlsx Sheet1 A1 'Hello World'" "Set A1"
run_test "cell set number" "$XLEX cell set $TEST_DIR/test.xlsx Sheet1 B1 100" "Set B1"
run_test "cell set number 2" "$XLEX cell set $TEST_DIR/test.xlsx Sheet1 C1 200" "Set C1"
run_test "cell get" "$XLEX cell get $TEST_DIR/test.xlsx Sheet1 A1" "Hello World"
run_test "cell formula (with =)" "$XLEX cell formula $TEST_DIR/test.xlsx Sheet1 D1 '=B1+C1'" "=B1+C1"
run_test "cell formula (without =)" "$XLEX cell formula $TEST_DIR/test.xlsx Sheet1 E1 'B1*C1'" "=B1*C1"
run_test "cell type" "$XLEX cell type $TEST_DIR/test.xlsx Sheet1 A1" "string"
run_test "cell type number" "$XLEX cell type $TEST_DIR/test.xlsx Sheet1 B1" "number"
run_test "cell clear" "$XLEX cell clear $TEST_DIR/test.xlsx Sheet1 E1" "Cleared"

# Verify formula doesn't have double =
FORMULA_VALUE=$($XLEX cell get $TEST_DIR/test.xlsx Sheet1 D1)
if [[ "$FORMULA_VALUE" == "=B1+C1" ]] && [[ "$FORMULA_VALUE" != "==B1+C1" ]]; then
    log_pass "formula no double equals"
else
    log_fail "formula no double equals - Got: $FORMULA_VALUE"
fi

# ===== Phase 5: Row Operations =====
echo ""
echo "--- Phase 5: Row Operations ---"

run_test "row get" "$XLEX row get $TEST_DIR/test.xlsx Sheet1 1" "Hello World"
run_test "row append" "$XLEX row append $TEST_DIR/test.xlsx Sheet1 'Name,Age,City'" "Appended"
run_test "row insert" "$XLEX row insert $TEST_DIR/test.xlsx Sheet1 3" "Inserted"
run_test "row delete" "$XLEX row delete $TEST_DIR/test.xlsx Sheet1 3" "Deleted"
run_test "row copy" "$XLEX row copy $TEST_DIR/test.xlsx Sheet1 1 5" "Copied"
run_test "row move" "$XLEX row move $TEST_DIR/test.xlsx Sheet1 5 6" "Moved"
run_test "row height" "$XLEX row height $TEST_DIR/test.xlsx Sheet1 1 30" "Set row 1 height"
run_test "row hide" "$XLEX row hide $TEST_DIR/test.xlsx Sheet1 2" "Hid row"
run_test "row unhide" "$XLEX row unhide $TEST_DIR/test.xlsx Sheet1 2" "Unhid row"
run_test "row find" "$XLEX row find $TEST_DIR/test.xlsx Sheet1 'Hello'" "1"

# ===== Phase 6: Column Operations =====
echo ""
echo "--- Phase 6: Column Operations ---"

run_test "column get" "$XLEX column get $TEST_DIR/test.xlsx Sheet1 A" "Hello World"
run_test "column insert" "$XLEX column insert $TEST_DIR/test.xlsx Sheet1 B" "Inserted"
run_test "column delete" "$XLEX column delete $TEST_DIR/test.xlsx Sheet1 B" "Deleted"
run_test "column copy" "$XLEX column copy $TEST_DIR/test.xlsx Sheet1 A D" "Copied"
run_test "column move" "$XLEX column move $TEST_DIR/test.xlsx Sheet1 D E" "Moved"
run_test "column width" "$XLEX column width $TEST_DIR/test.xlsx Sheet1 A 20" "Set column A width"
run_test "column hide" "$XLEX column hide $TEST_DIR/test.xlsx Sheet1 B" "Hid column"
run_test "column unhide" "$XLEX column unhide $TEST_DIR/test.xlsx Sheet1 B" "Unhid column"
run_test "column header" "$XLEX column header $TEST_DIR/test.xlsx Sheet1 A" "Hello World"
run_test "column find" "$XLEX column find $TEST_DIR/test.xlsx Sheet1 'Hello'" "A"

# ===== Phase 7: Range Operations =====
echo ""
echo "--- Phase 7: Range Operations ---"

run_test "range get" "$XLEX range get $TEST_DIR/test.xlsx Sheet1 A1:D2" "Row"
run_test "range fill" "$XLEX range fill $TEST_DIR/test.xlsx Sheet1 A10:C10 'X,Y,Z'" "Filled"
run_test "range copy" "$XLEX range copy $TEST_DIR/test.xlsx Sheet1 A1:B2 A20" "Copied"
run_test "range move" "$XLEX range move $TEST_DIR/test.xlsx Sheet1 A20:B21 A25" "Moved"
run_test "range clear" "$XLEX range clear $TEST_DIR/test.xlsx Sheet1 A25:B26" "Cleared"
run_test "range merge" "$XLEX range merge $TEST_DIR/test.xlsx Sheet1 A30:C30" "Merged"
run_test "range unmerge" "$XLEX range unmerge $TEST_DIR/test.xlsx Sheet1 A30:C30" "Unmerged"
run_test "range validate" "$XLEX range validate $TEST_DIR/test.xlsx Sheet1 A1:A1 nonempty" ""
run_test "range sort" "$XLEX range sort $TEST_DIR/test.xlsx Sheet1 A10:C10" "Sorted"

# ===== Phase 8: Formula Operations =====
echo ""
echo "--- Phase 8: Formula Operations ---"

run_test "formula validate" "$XLEX formula validate '=SUM(A1:A10)'" "valid"
run_test "formula list" "$XLEX formula list $TEST_DIR/test.xlsx Sheet1" "formula"
run_test "formula stats" "$XLEX formula stats $TEST_DIR/test.xlsx Sheet1" "Total formulas:"
run_test "formula refs" "$XLEX formula refs $TEST_DIR/test.xlsx Sheet1 D1" ""
run_test "formula replace" "$XLEX formula replace $TEST_DIR/test.xlsx Sheet1 'B1' 'B2'" ""

# ===== Phase 9: Style Operations =====
echo ""
echo "--- Phase 9: Style Operations ---"

run_test "style list" "$XLEX style list $TEST_DIR/test.xlsx" "Styles:"
run_test "style preset list" "$XLEX style preset list" "header"

# ===== Phase 10: Import/Export Operations =====
echo ""
echo "--- Phase 10: Import/Export Operations ---"

# Create test CSV
echo 'Name,Age,City
Alice,25,Taipei
Bob,30,Tokyo
Carol,28,Seoul' > "$TEST_DIR/data.csv"

run_test "import csv" "$XLEX import csv $TEST_DIR/data.csv $TEST_DIR/imported.xlsx" "Imported"
run_test "export csv" "$XLEX export csv $TEST_DIR/imported.xlsx - --sheet Sheet1" "Name"
run_test "export json" "$XLEX export json $TEST_DIR/imported.xlsx -" "["
run_test "export ndjson" "$XLEX export ndjson $TEST_DIR/imported.xlsx -" "Name"
run_test "export meta" "$XLEX export meta $TEST_DIR/test.xlsx $TEST_DIR/meta.json" "Exported metadata"
run_test "convert" "$XLEX convert $TEST_DIR/test.xlsx $TEST_DIR/converted.xlsx" "Copied"

# Check export file
if [[ -f "$TEST_DIR/meta.json" ]]; then
    log_pass "meta.json file created"
else
    log_fail "meta.json file not created"
fi

# ===== Phase 11: Template Operations =====
echo ""
echo "--- Phase 11: Template Operations ---"

run_test "template validate" "$XLEX template validate $TEST_DIR/test.xlsx" ""

# ===== Phase 12: Utility Operations =====
echo ""
echo "--- Phase 12: Utility Operations ---"

run_test "completion bash" "$XLEX completion bash" "complete"
run_test "completion zsh" "$XLEX completion zsh" "compdef"
run_test "completion fish" "$XLEX completion fish" "complete"

# ===== Phase 13: JSON Format Output =====
echo ""
echo "--- Phase 13: JSON Format Output ---"

run_test "info json" "$XLEX info $TEST_DIR/test.xlsx --format json" '"file":'
run_test "stats json" "$XLEX stats $TEST_DIR/test.xlsx --format json" '"sheet_count":'

# ===== Phase 14: Error Handling =====
echo ""
echo "--- Phase 14: Error Handling ---"

run_test_fail "open nonexistent file" "$XLEX info /nonexistent/file.xlsx"
run_test_fail "invalid cell reference" "$XLEX cell get $TEST_DIR/test.xlsx Sheet1 INVALID"
run_test_fail "nonexistent sheet" "$XLEX cell get $TEST_DIR/test.xlsx NoSuchSheet A1"

# ===== Summary =====
echo ""
echo "========================================"
echo "              Test Summary"
echo "========================================"
echo -e "${GREEN}Passed:${NC}  $PASSED"
echo -e "${RED}Failed:${NC}  $FAILED"
echo -e "${YELLOW}Skipped:${NC} $SKIPPED"
echo "----------------------------------------"
TOTAL=$((PASSED + FAILED + SKIPPED))
echo "Total:   $TOTAL"
echo ""

if [[ $FAILED -eq 0 ]]; then
    echo -e "${GREEN}All tests passed!${NC}"
    exit 0
else
    echo -e "${RED}Some tests failed!${NC}"
    exit 1
fi
