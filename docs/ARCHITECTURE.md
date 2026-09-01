# Architecture

Excel Diff Tracker is a local, per-user Windows tray application built as separate testable projects.

## Components

- `ExcelDiffTracker.Core` reads guarded Open XML packages into canonical workbook snapshots and computes exact semantic deltas.
- `ExcelDiffTracker.Storage` owns versioned SQLite migrations, baselines, versions, change rows, errors, settings, and the pending-report outbox.
- `ExcelDiffTracker.Reporting` renders Markdown from committed version data using atomic file replacement.
- `ExcelDiffTracker.Tracking` combines file-system notifications, periodic reconciliation, per-workbook workers, stability checks, temporary read-only copies, SHA-256 deduplication, retries, and report recovery.
- `ExcelDiffTracker.App` is the WPF shell, onboarding flow, tray integration, theme system, and user settings.

## Save flow

1. A watcher event or reconciliation marks one workbook dirty.
2. Its dedicated worker waits briefly for the save burst to settle; it never cancels an in-flight accepted candidate.
3. The worker opens a stable read-only candidate while excluding concurrent writers and copies it to a private temporary file.
4. The package guard checks type, paths, ZIP expansion limits, and package structure.
5. The file is hashed. A hash already represented by the current baseline is ignored as a duplicate observation.
6. The Open XML snapshot is extracted and compared with the current canonical snapshot.
7. One SQLite transaction writes the version, all semantic deltas, the new snapshot, and a pending report record.
8. Markdown is written atomically and the version is marked ready. Pending reports are regenerated after interruption.

The baseline advances only after extraction and the SQLite transaction succeed. Corrupt, encrypted, unsafe, or unsupported workbooks create an error record without changing it.

## Formula semantics

Formula XML is compared exactly as stored, as required by the product definition. Cached formula results are compared separately. Excel Diff Tracker does not calculate formulas or decide whether two formulas are mathematically equivalent.

## Storage model

The SQLite database is `%LocalAppData%\Excel Diff Tracker\history.db`. It contains the current semantic snapshot and all deltas, not historical workbook files. WAL mode and foreign keys are enabled. The schema version is migrated transactionally and a database from a newer unsupported schema is rejected.

## Security boundary

Macros are never executed. Excel automation is not used by the installed application. The optional developer smoke test uses a real Excel instance solely to create and save test workbooks, then verifies the resulting tracker data.
