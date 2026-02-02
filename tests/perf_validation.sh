#!/bin/bash
# Performance validation script for XLEX CLI
# Run with: ./tests/perf_validation.sh

set -e

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
PROJECT_DIR="$(dirname "$SCRIPT_DIR")"
FIXTURES_DIR="$SCRIPT_DIR/fixtures"
TEMP_DIR=$(mktemp -d)
XLEX="cargo run --release --bin xlex --"

# Colors for output
RED='\033[0;31m'
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
NC='\033[0m' # No Color

# Performance targets (in milliseconds)
TARGET_SHEET_LIST=100
TARGET_COLUMN_READ=300
TARGET_CELL_UPDATE=200
TARGET_10K_APPEND=1000

cleanup() {
    rm -rf "$TEMP_DIR"
}
trap cleanup EXIT

echo "=========================================="
echo "XLEX Performance Validation"
echo "=========================================="
echo ""

# Build release version first
echo "Building release version..."
cd "$PROJECT_DIR"
cargo build --release --bin xlex 2>/dev/null

XLEX_BIN="$PROJECT_DIR/target/release/xlex"

if [ ! -f "$XLEX_BIN" ]; then
    echo -e "${RED}Error: xlex binary not found${NC}"
    exit 1
fi

# Function to measure command execution time
measure_time() {
    local start end duration_ms
    start=$(date +%s%N)
    "$@" > /dev/null 2>&1
    end=$(date +%s%N)
    duration_ms=$(( (end - start) / 1000000 ))
    echo $duration_ms
}

# Function to check if target met
check_target() {
    local name=$1
    local actual=$2
    local target=$3
    
    if [ "$actual" -le "$target" ]; then
        echo -e "${GREEN}✓${NC} $name: ${actual}ms (target: <${target}ms)"
        return 0
    else
        echo -e "${RED}✗${NC} $name: ${actual}ms (target: <${target}ms) - EXCEEDED"
        return 1
    fi
}

echo ""
echo "Creating test fixtures..."

# Create a medium-sized workbook (10K rows)
MEDIUM_FILE="$TEMP_DIR/medium.xlsx"
"$XLEX_BIN" create "$MEDIUM_FILE"
for i in $(seq 1 100); do
    "$XLEX_BIN" cell set "$MEDIUM_FILE" Sheet1 "A$i" "Value $i" 2>/dev/null
done

# Create a larger workbook for append test
LARGE_FILE="$TEMP_DIR/large.xlsx"
"$XLEX_BIN" create "$LARGE_FILE"

echo "Test fixtures created."
echo ""
echo "=========================================="
echo "Running Performance Tests"
echo "=========================================="
echo ""

FAILURES=0

# Test 1: Sheet list (target <100ms)
echo "Test 1: Sheet List"
duration=$(measure_time "$XLEX_BIN" sheet list "$MEDIUM_FILE")
check_target "sheet list" "$duration" "$TARGET_SHEET_LIST" || ((FAILURES++))

# Test 2: Column read (target <300ms)
echo "Test 2: Column Read"
duration=$(measure_time "$XLEX_BIN" column get "$MEDIUM_FILE" Sheet1 A)
check_target "column get A" "$duration" "$TARGET_COLUMN_READ" || ((FAILURES++))

# Test 3: Cell update (target <200ms)
echo "Test 3: Cell Update"
duration=$(measure_time "$XLEX_BIN" cell set "$MEDIUM_FILE" Sheet1 A50 "Updated")
check_target "cell set" "$duration" "$TARGET_CELL_UPDATE" || ((FAILURES++))

# Test 4: Row append performance (target <1s for 10K rows conceptually)
# Note: We test 100 rows here as a proxy since actual 10K would be slow to setup
echo "Test 4: Batch Row Operations"
BATCH_FILE="$TEMP_DIR/batch.xlsx"
"$XLEX_BIN" create "$BATCH_FILE"
start=$(date +%s%N)
for i in $(seq 1 100); do
    "$XLEX_BIN" row append "$BATCH_FILE" Sheet1 "Val$i" 2>/dev/null
done
end=$(date +%s%N)
duration_ms=$(( (end - start) / 1000000 ))
# Scale: if 100 rows takes X ms, 10K would take ~100X ms
# Target: 100 rows should take <10ms to meet 1s for 10K target
scaled_target=$((TARGET_10K_APPEND / 100))
check_target "100 row append (scaled)" "$duration_ms" "$scaled_target" || {
    echo -e "${YELLOW}  Note: Actual 10K row append would take ~$((duration_ms * 100))ms${NC}"
    ((FAILURES++))
}

# Test 5: Export to CSV
echo "Test 5: Export to CSV"
CSV_FILE="$TEMP_DIR/export.csv"
duration=$(measure_time "$XLEX_BIN" export csv "$MEDIUM_FILE" "$CSV_FILE")
check_target "export csv (100 rows)" "$duration" "500" || ((FAILURES++))

# Test 6: Workbook info
echo "Test 6: Workbook Info"
duration=$(measure_time "$XLEX_BIN" info "$MEDIUM_FILE")
check_target "workbook info" "$duration" "100" || ((FAILURES++))

# Test 7: Workbook validation
echo "Test 7: Workbook Validation"
duration=$(measure_time "$XLEX_BIN" validate "$MEDIUM_FILE")
check_target "workbook validate" "$duration" "200" || ((FAILURES++))

echo ""
echo "=========================================="
echo "Memory Usage Test"
echo "=========================================="
echo ""

# Check memory usage if available
if command -v /usr/bin/time &> /dev/null; then
    echo "Measuring peak memory for opening workbook..."
    MEM_OUTPUT=$(/usr/bin/time -v "$XLEX_BIN" info "$MEDIUM_FILE" 2>&1 | grep "Maximum resident set size" || echo "N/A")
    echo "Memory: $MEM_OUTPUT"
else
    echo "Note: /usr/bin/time not available for memory measurement"
fi

echo ""
echo "=========================================="
echo "Summary"
echo "=========================================="
echo ""

if [ $FAILURES -eq 0 ]; then
    echo -e "${GREEN}All performance targets met!${NC}"
    exit 0
else
    echo -e "${RED}$FAILURES performance target(s) not met${NC}"
    echo ""
    echo "Note: Performance may vary based on system load and hardware."
    echo "Consider running on a quiescent system for accurate results."
    exit 1
fi
