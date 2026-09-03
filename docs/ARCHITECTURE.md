# Architecture

Excel Scenario Analysis Tool is a local, per-user Windows tray application built as separate testable projects. The existing `ExcelDiffTracker.*` namespaces remain internal implementation names in 0.2.0.

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

## Release payload identity

The application and external AcceptanceProbe publish as self-contained single-file ARM64 executables, including their native libraries. Each frozen executable SHA-256 therefore covers its managed code, dependencies, and runtime. The release build and installed gates reject loose DLLs or runtime configuration files alongside either executable. In-place upgrade removes the explicitly enumerated runtime files shipped by 0.1.1 from the application directory; it does not target user workbooks, history, reports, or settings. The SDK's matching build metadata supplies the exact third-party license versions; it is embedded in the delivered executable.

Native libraries are extracted into the Windows user's temporary .NET bundle cache at startup. This standard [.NET single-file deployment behavior](https://learn.microsoft.com/en-us/dotnet/core/deploying/single-file/overview) requires separate installed startup and SQLite acceptance. Local supporting tests direct extraction into disposable test output.

## Formula semantics

Formula XML is compared exactly as stored, as required by the product definition. Cached formula results are compared separately. Excel Scenario Analysis Tool does not calculate formulas or decide whether two formulas are mathematically equivalent.

## Storage model

The 0.2 SQLite database is `%LocalAppData%\Excel Scenario Analysis Tool\history.db`. It contains the current semantic snapshot and all deltas, not historical workbook files. WAL mode and foreign keys are enabled. The schema version is migrated transactionally and a database from a newer unsupported schema is rejected.

## Comparison views

Every accepted save remains a forward delta from the immediately preceding scan and every automatic Markdown report remains chronological. A per-workbook setting selects the comparison view: previous save, original scan 0, or a specific saved scan identified by its stable database ID. Historical snapshots are reconstructed locally by reverse-applying later sheet and cell deltas to the current semantic snapshot. The normal comparator then computes the selected-baseline view. Explicit derived exports identify both scan numbers and hashes and state that they do not replace chronological evidence.

## Security boundary

Macros are never executed. Excel automation is not used by the installed application. The optional developer smoke test uses a real Excel instance solely to create and save test workbooks, then verifies the resulting tracker data.
