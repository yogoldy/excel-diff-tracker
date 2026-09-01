using ExcelDiffTracker.Core;

namespace ExcelDiffTracker.Tracking;

public enum CaptureEventKind
{
    BaselineCreated,
    Captured,
    NoTrackedChanges,
    DuplicateIgnored,
    Failed,
    StatusChanged
}

public sealed record CaptureEvent
{
    public required Guid WorkbookId { get; init; }
    public required string WorkbookPath { get; init; }
    public required CaptureEventKind Kind { get; init; }
    public required DateTime OccurredUtc { get; init; }
    public string? Message { get; init; }
    public WorkbookDiff? Diff { get; init; }
    public string? ReportPath { get; init; }
    public Exception? Exception { get; init; }
}
