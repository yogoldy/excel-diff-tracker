using System.Collections.ObjectModel;

namespace ExcelDiffTracker.Core;

public sealed record CellState
{
    public required uint SheetId { get; init; }
    public required string SheetName { get; init; }
    public required string Address { get; init; }
    public required string CellType { get; init; }
    public string? LiteralValue { get; init; }
    public string? FormulaText { get; init; }
    public string? FormulaType { get; init; }
    public string? FormulaReference { get; init; }
    public uint? FormulaSharedIndex { get; init; }
    public string? FormulaXml { get; init; }
    public string? CachedResult { get; init; }

    public bool HasFormula => FormulaText is not null || FormulaType is not null || FormulaSharedIndex is not null;

    public string FormulaIdentity => FormulaXml ?? string.Join(
        "\u001f",
        FormulaText ?? string.Empty,
        FormulaType ?? string.Empty,
        FormulaReference ?? string.Empty,
        FormulaSharedIndex?.ToString(System.Globalization.CultureInfo.InvariantCulture) ?? string.Empty);
}

public sealed record SheetState
{
    public required uint SheetId { get; init; }
    public required string Name { get; init; }
    public required int Position { get; init; }
    public required string Visibility { get; init; }
    public IReadOnlyDictionary<string, CellState> Cells { get; init; } =
        new ReadOnlyDictionary<string, CellState>(new Dictionary<string, CellState>());
}

public sealed record WorkbookSnapshot
{
    public required string SourcePath { get; init; }
    public required DateTime CapturedAtUtc { get; init; }
    public IReadOnlyDictionary<uint, SheetState> Sheets { get; init; } =
        new ReadOnlyDictionary<uint, SheetState>(new Dictionary<uint, SheetState>());

    public int PopulatedCellCount => Sheets.Values.Sum(sheet => sheet.Cells.Count);
}

public enum CellChangeKind
{
    LiteralAdded,
    LiteralChanged,
    LiteralCleared,
    FormulaAdded,
    FormulaChanged,
    FormulaRemoved,
    FormulaResultChanged,
    CellTypeChanged
}

public enum SheetChangeKind
{
    Added,
    Removed,
    Renamed,
    Reordered,
    VisibilityChanged
}

public sealed record CellDelta
{
    public required uint SheetId { get; init; }
    public required string SheetName { get; init; }
    public required string Address { get; init; }
    public required IReadOnlyList<CellChangeKind> Kinds { get; init; }
    public CellState? Before { get; init; }
    public CellState? After { get; init; }
}

public sealed record SheetDelta
{
    public required uint SheetId { get; init; }
    public required SheetChangeKind Kind { get; init; }
    public SheetState? Before { get; init; }
    public SheetState? After { get; init; }
}

public sealed record WorkbookDiff
{
    public required IReadOnlyList<CellDelta> CellChanges { get; init; }
    public required IReadOnlyList<SheetDelta> SheetChanges { get; init; }

    public bool HasTrackedChanges => CellChanges.Count != 0 || SheetChanges.Count != 0;
    public int FormulaChangeCount => CellChanges.Count(change =>
        change.Kinds.Any(kind => kind is CellChangeKind.FormulaAdded or CellChangeKind.FormulaChanged or CellChangeKind.FormulaRemoved));
    public int FormulaResultChangeCount => CellChanges.Count(change => change.Kinds.Contains(CellChangeKind.FormulaResultChanged));
    public int LiteralChangeCount => CellChanges.Count(change =>
        change.Kinds.Any(kind => kind is CellChangeKind.LiteralAdded or CellChangeKind.LiteralChanged or CellChangeKind.LiteralCleared));
    public int CellTypeChangeCount => CellChanges.Count(change => change.Kinds.Contains(CellChangeKind.CellTypeChanged));
}

public enum TrackingStatus
{
    Active,
    Processing,
    Paused,
    Missing,
    Warning
}

public sealed record TrackedWorkbook
{
    public required Guid Id { get; init; }
    public required string Path { get; init; }
    public required string ReportDirectory { get; init; }
    public required bool IsEnabled { get; init; }
    public required TrackingStatus Status { get; init; }
    public required DateTime CreatedUtc { get; init; }
    public DateTime? LastSuccessfulCaptureUtc { get; init; }
    public required long CurrentSequence { get; init; }
    public required string CurrentHash { get; init; }
    public string? LastSummary { get; init; }
    public string? LastError { get; init; }
}

public enum VersionStatus
{
    Baseline,
    Captured,
    Error
}

public enum ReportStatus
{
    Pending,
    Ready
}

public sealed record VersionRecord
{
    public required long Id { get; init; }
    public required Guid WorkbookId { get; init; }
    public required string WorkbookPath { get; init; }
    public required long Sequence { get; init; }
    public required DateTime CapturedUtc { get; init; }
    public required DateTime FileLastWriteUtc { get; init; }
    public required string PreviousSha256 { get; init; }
    public required string Sha256 { get; init; }
    public required VersionStatus Status { get; init; }
    public required ReportStatus ReportStatus { get; init; }
    public required int CellChangeCount { get; init; }
    public required int SheetChangeCount { get; init; }
    public required int LiteralChangeCount { get; init; }
    public required int FormulaChangeCount { get; init; }
    public required int FormulaResultChangeCount { get; init; }
    public required int CellTypeChangeCount { get; init; }
    public string? ReportPath { get; init; }
    public string? Summary { get; init; }
}

public sealed record CaptureErrorRecord
{
    public required long Id { get; init; }
    public required Guid WorkbookId { get; init; }
    public required string WorkbookPath { get; init; }
    public required DateTime FirstSeenUtc { get; init; }
    public required DateTime LastSeenUtc { get; init; }
    public required int OccurrenceCount { get; init; }
    public string? Sha256 { get; init; }
    public required string Category { get; init; }
    public required string Message { get; init; }
}

public sealed class UnsupportedWorkbookException(string message) : Exception(message);
public sealed class UnsafeWorkbookException(string message) : Exception(message);
