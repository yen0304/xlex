# Exit Codes

xlex uses standard exit codes to indicate the result of command execution.

## Exit Code Summary

| Code | Name | Description |
|------|------|-------------|
| `0` | `SUCCESS` | Command completed successfully |
| `1` | `ERROR_GENERAL` | General/unknown error |
| `2` | `ERROR_ARGUMENTS` | Invalid command-line arguments |
| `3` | `ERROR_FILE_NOT_FOUND` | Input file not found |
| `4` | `ERROR_PERMISSION` | Permission denied |
| `5` | `ERROR_FORMAT` | Invalid file format |
| `6` | `ERROR_REFERENCE` | Invalid cell/range reference |
| `7` | `ERROR_SHEET` | Sheet not found or invalid |
| `8` | `ERROR_FORMULA` | Formula error |
| `9` | `ERROR_TEMPLATE` | Template processing error |
| `10` | `ERROR_CONFIG` | Configuration error |
| `11` | `ERROR_IO` | I/O error (read/write failure) |
| `12` | `ERROR_VALIDATION` | Validation failed |
| `13` | `ERROR_LIMIT` | Resource limit exceeded |
| `14` | `ERROR_NETWORK` | Network error |
| `15` | `ERROR_TIMEOUT` | Operation timed out |

## Detailed Descriptions

### 0 - SUCCESS

The command completed successfully without errors.

```bash
xlex info report.xlsx
echo $?  # 0
```

### 1 - ERROR_GENERAL

A general error occurred that doesn't fit other categories.

```bash
xlex some-invalid-command
echo $?  # 1
```

### 2 - ERROR_ARGUMENTS

Invalid or missing command-line arguments.

**Common causes:**
- Missing required arguments
- Invalid option values
- Conflicting options

```bash
xlex cell set report.xlsx  # Missing cell reference and value
echo $?  # 2
```

### 3 - ERROR_FILE_NOT_FOUND

The specified file does not exist.

```bash
xlex info nonexistent.xlsx
echo $?  # 3
```

### 4 - ERROR_PERMISSION

Permission denied when accessing a file.

**Common causes:**
- File is read-only
- Insufficient filesystem permissions
- File locked by another process

```bash
xlex cell set readonly.xlsx A1 "value"
echo $?  # 4
```

### 5 - ERROR_FORMAT

The file format is invalid or corrupted.

**Common causes:**
- File is not a valid Excel file
- File is corrupted
- Unsupported Excel version

```bash
xlex info notexcel.txt
echo $?  # 5
```

### 6 - ERROR_REFERENCE

Invalid cell or range reference.

**Common causes:**
- Invalid cell notation (e.g., `AA` instead of `A1`)
- Range outside valid bounds
- Invalid named range

```bash
xlex cell get report.xlsx InvalidRef
echo $?  # 6
```

### 7 - ERROR_SHEET

Sheet-related error.

**Common causes:**
- Sheet does not exist
- Cannot delete last sheet
- Duplicate sheet name

```bash
xlex sheet remove report.xlsx "NonexistentSheet"
echo $?  # 7
```

### 8 - ERROR_FORMULA

Formula-related error.

**Common causes:**
- Invalid formula syntax
- Circular reference detected
- Function not supported

```bash
xlex formula validate "=SUM(A1:A10"  # Missing closing paren
echo $?  # 8
```

### 9 - ERROR_TEMPLATE

Template processing error.

**Common causes:**
- Invalid template syntax
- Missing required variables
- Template/data mismatch

```bash
xlex template apply invalid-template.xlsx output.xlsx --data data.json
echo $?  # 9
```

### 10 - ERROR_CONFIG

Configuration error.

**Common causes:**
- Invalid configuration file
- Unknown configuration key
- Invalid configuration value

```bash
xlex config validate invalid-config.yml
echo $?  # 10
```

### 11 - ERROR_IO

I/O error during file operations.

**Common causes:**
- Disk full
- Network drive disconnected
- File system error

```bash
xlex clone report.xlsx /readonly/output.xlsx
echo $?  # 11
```

### 12 - ERROR_VALIDATION

Data validation failed.

**Common causes:**
- Workbook validation errors
- Data validation rules violated
- Schema validation failure

```bash
xlex validate corrupted.xlsx --strict
echo $?  # 12
```

### 13 - ERROR_LIMIT

Resource limit exceeded.

**Common causes:**
- File too large
- Too many operations
- Memory limit exceeded

```bash
xlex from csv huge.csv output.xlsx  # File exceeds limits
echo $?  # 13
```

### 14 - ERROR_NETWORK

Network-related error.

**Common causes:**
- Connection timeout
- DNS resolution failure
- Remote server unavailable

```bash
xlex template apply template.xlsx out.xlsx --data "https://unavailable.example.com/data"
echo $?  # 14
```

### 15 - ERROR_TIMEOUT

Operation timed out.

**Common causes:**
- Operation took too long
- External service timeout
- Lock acquisition timeout

```bash
xlex batch long-running.txt --timeout 60
echo $?  # 15
```

## Using Exit Codes in Scripts

### Bash

```bash
#!/bin/bash

xlex validate report.xlsx
case $? in
  0)
    echo "File is valid"
    ;;
  5)
    echo "File format is invalid"
    ;;
  12)
    echo "Validation errors found"
    ;;
  *)
    echo "Unknown error"
    ;;
esac
```

### Conditional Execution

```bash
# Only proceed if file is valid
xlex validate data.xlsx && xlex to csv data.xlsx > export.csv

# Handle errors
xlex info report.xlsx || echo "Failed to read file"

# Check specific exit codes
if ! xlex sheet remove report.xlsx "TempSheet" 2>/dev/null; then
  if [ $? -eq 7 ]; then
    echo "Sheet doesn't exist, continuing..."
  else
    echo "Unexpected error"
    exit 1
  fi
fi
```

### Error Handling in CI/CD

```yaml
# GitHub Actions example
- name: Validate Excel files
  run: |
    for file in data/*.xlsx; do
      if ! xlex validate "$file"; then
        echo "::error file=$file::Validation failed"
        exit 1
      fi
    done
```

## JSON Error Output

Use `--json-errors` for machine-readable error output:

```bash
xlex info nonexistent.xlsx --json-errors
```

```json
{
  "success": false,
  "exit_code": 3,
  "error": {
    "code": "XLEX_E010",
    "message": "File not found: nonexistent.xlsx",
    "suggestion": "Check if the file path is correct"
  }
}
```

## See Also

- [Error Codes](error-codes.md) - Detailed error code reference
- [CLI Reference](cli-reference.md) - Complete CLI documentation
