# Security policy

## Supported versions

Security fixes are provided for the latest published release.

## Reporting a vulnerability

Please use the repository’s private **Report a vulnerability** feature under the Security tab. Do not include workbook contents or other private data in a public issue.

Excel Diff Tracker processes files locally and intentionally has no network client or telemetry. A rejected or malformed workbook must never advance its saved baseline. Reports from untrusted workbooks should still be treated as untrusted text when opened in a Markdown renderer.
