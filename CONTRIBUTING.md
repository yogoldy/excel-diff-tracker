# Contributing

Contributions are welcome. Please open an issue before a large change so the intended behavior can be agreed first.

## Development setup

- Windows 11
- .NET 10 SDK
- PowerShell 7 or Windows PowerShell
- Inno Setup 7 only when building the installer
- Excel only for the optional live-save smoke test

No Python or virtual environment is used.

Before submitting a change, run:

```powershell
dotnet build .\ExcelDiffTracker.slnx -c Release
dotnet test .\ExcelDiffTracker.slnx -c Release --no-build
```

Keep the diff engine, tracking pipeline, storage, reporting, and WPF interface separated. Add tests for semantic behavior and failure paths. Do not add telemetry, cloud processing, macro execution, or a dependency on an installed copy of Excel.

By contributing, you agree that your contribution is licensed under the repository’s MIT License.
