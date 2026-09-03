# Privacy

Excel Scenario Analysis Tool is private by design.

- Workbook processing happens entirely on the Windows computer where the app is installed.
- No workbook content, filename, path, hash, history, report, crash detail, or usage event is transmitted.
- The app contains no telemetry, analytics, advertising, account system, AI service, web API, or update service.
- New 0.2 SQLite history is stored at `%LocalAppData%\Excel Scenario Analysis Tool\history.db`.
- Markdown reports are stored only in folders selected by the user.
- Temporary workbook copies are private local processing files and are normally removed after capture. An operating-system or storage failure can leave a local temporary copy until later cleanup.

Pausing or removing tracking does not silently delete history. Permanent deletion requires a separate confirmation. On first 0.2 launch, legacy cleanup can delete only `history.db`, `history.db-shm`, and `history.db-wal` from the old local-app-data directory, and only after explicit confirmation. It never deletes Markdown reports or workbooks. Uninstalling the app preserves history and reports to avoid destructive surprise; use the confirmed purge action first if permanent removal is desired.
