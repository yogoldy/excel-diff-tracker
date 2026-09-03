# Excel Scenario Analysis Tool 0.2.0 — local candidate

This is an unqualified local candidate. It is not authorization to publish an installer or represent the application as production-ready.

## Added

- Original scenario-branches product identity with separate large and small-format icon treatments.
- Larger shared typography, responsive dashboard rows, and compact or expanded workbook and history layouts.
- Three-dot workbook and saved-scan menus with copy-path and File Explorer actions.
- Per-workbook comparison baselines: previous save, original scan 0, or a specific saved scan.
- On-screen selected-comparison summaries and explicit derived Markdown exports labeled with source and target scans and hashes.
- Native Windows 11 caption colors synchronized with light, dark, system, and high-contrast behavior.
- Runtime version and platform details on About, plus a deliberately disabled local-only donation control.

## Evidence guarantees

- Every accepted save remains an immutable chronological delta from the immediately preceding scan.
- Selecting a comparison baseline never rewrites automatic reports, scan hashes, or history rows.
- A derived comparison export is rejected if its destination is any automatic chronological report path.
- Historical comparison snapshots are reconstructed from the current snapshot and retained deltas; full workbook copies are not duplicated.
- The 0.2 app uses a fresh local database. Optional legacy cleanup targets only the prior SQLite database files and never Markdown reports or workbooks.

## Deferred

- Scenario recording sessions, named scenario libraries, an Excel add-in, and downstream visualization integration.
- Public-release qualification and publication.
