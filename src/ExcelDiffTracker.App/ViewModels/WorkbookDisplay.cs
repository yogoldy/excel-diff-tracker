using ExcelDiffTracker.Core;

namespace ExcelDiffTracker.App.ViewModels;

public sealed record WorkbookDisplay
{
    public required TrackedWorkbook Workbook { get; init; }
    public required string ComparisonBaselineLabel { get; init; }
    public required string ComparisonSummary { get; init; }

    public Guid Id => Workbook.Id;
    public string Path => Workbook.Path;
    public string FileName => System.IO.Path.GetFileName(Workbook.Path);
    public string Directory => System.IO.Path.GetDirectoryName(Workbook.Path) ?? Workbook.Path;
    public string ReportDirectory => Workbook.ReportDirectory;
    public bool IsEnabled => Workbook.IsEnabled;
    public TrackingStatus Status => Workbook.Status;
    public DateTime? LastSuccessfulCaptureUtc => Workbook.LastSuccessfulCaptureUtc;
    public long CurrentSequence => Workbook.CurrentSequence;
    public string? LastSummary => Workbook.LastSummary;
    public string? LastError => Workbook.LastError;
}
