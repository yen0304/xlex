# Security Policy

## Supported Versions

| Version | Supported          |
| ------- | ------------------ |
| 0.1.x   | :white_check_mark: |

## Reporting a Vulnerability

If you discover a security vulnerability in xlex, please report it by:

1. Opening a [private security advisory](https://github.com/yen0304/xlex/security/advisories/new)
2. Or contacting [@yen0304](https://github.com/yen0304) directly via GitHub

Please include:

1. Description of the vulnerability
2. Steps to reproduce
3. Potential impact
4. Suggested fix (if any)

## Response Timeline

- **Acknowledgment**: Within 48 hours
- **Initial Assessment**: Within 7 days
- **Fix Timeline**: Depends on severity
  - Critical: 24-48 hours
  - High: 7 days
  - Medium: 30 days
  - Low: Next release

## Disclosure Policy

- We follow responsible disclosure
- We will credit reporters (unless anonymity is requested)
- Public disclosure after fix is released

## Security Best Practices

When using xlex:

1. **Validate input files** - Don't process untrusted xlsx files without validation
2. **Use latest version** - Keep xlex updated
3. **Review templates** - Template placeholders can execute data transformations
4. **Limit permissions** - Run with minimal file system permissions

## Known Limitations

- xlex does not execute Excel macros (VBA)
- xlex does not evaluate formulas (values only)
- Template processing is sandboxed (no arbitrary code execution)
