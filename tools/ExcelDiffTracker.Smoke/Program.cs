using System.Text.Json;
using ExcelDiffTracker.Storage;
using ExcelDiffTracker.Tracking;

if (args is ["--extract", var extractPath])
{
    var fullPath = Path.GetFullPath(extractPath);
    var hashBefore = Convert.ToHexString(System.Security.Cryptography.SHA256.HashData(await File.ReadAllBytesAsync(fullPath)));
    var stopwatch = System.Diagnostics.Stopwatch.StartNew();
    var snapshot = new ExcelDiffTracker.Core.WorkbookExtractor().Extract(fullPath);
    stopwatch.Stop();
    var hashAfter = Convert.ToHexString(System.Security.Cryptography.SHA256.HashData(await File.ReadAllBytesAsync(fullPath)));
    var cellCount = snapshot.Sheets.Values.Sum(sheet => sheet.Cells.Count);
    var workingSetMiB = System.Diagnostics.Process.GetCurrentProcess().PeakWorkingSet64 / 1024d / 1024d;
    Console.WriteLine($"EXTRACT_PASS|sheets={snapshot.Sheets.Count}|cells={cellCount}|unchanged={hashBefore == hashAfter}|elapsedMs={stopwatch.ElapsedMilliseconds}|peakMiB={workingSetMiB:F1}");
    return hashBefore == hashAfter ? 0 : 1;
}

if (args is ["--expect-rejected", var rejectedPath])
{
    try
    {
        _ = new ExcelDiffTracker.Core.WorkbookExtractor().Extract(Path.GetFullPath(rejectedPath));
        Console.Error.WriteLine("REJECT_FAILED|workbook was accepted");
        return 1;
    }
    catch (Exception exception) when (exception is ExcelDiffTracker.Core.UnsafeWorkbookException or ExcelDiffTracker.Core.UnsupportedWorkbookException)
    {
        Console.WriteLine($"REJECT_PASS|{exception.GetType().Name}|{exception.Message}");
        return 0;
    }
}

if (args.Length != 5)
{
    Console.Error.WriteLine("Usage: ExcelDiffTracker.Smoke --extract <workbook>");
    Console.Error.WriteLine("   or: ExcelDiffTracker.Smoke --expect-rejected <workbook>");
    Console.Error.WriteLine("   or: ExcelDiffTracker.Smoke <workbook> <reports> <database> <ready-signal> <result-json>");
    return 2;
}

var workbookPath = Path.GetFullPath(args[0]);
var reportsPath = Path.GetFullPath(args[1]);
var databasePath = Path.GetFullPath(args[2]);
var readySignal = Path.GetFullPath(args[3]);
var resultPath = Path.GetFullPath(args[4]);
var store = new HistoryStore(databasePath);
await using var coordinator = new TrackingCoordinator(
    store,
    stableCopy: new StableWorkbookCopy(TimeSpan.FromMilliseconds(100), TimeSpan.FromMilliseconds(250), TimeSpan.FromSeconds(30)),
    reconciliationInterval: TimeSpan.FromSeconds(1));
await coordinator.InitializeAsync();
var tracked = await coordinator.AddWorkbookAsync(workbookPath, reportsPath);
var completion = new TaskCompletionSource<CaptureEvent>(TaskCreationOptions.RunContinuationsAsynchronously);
coordinator.CaptureOccurred += (_, captureEvent) =>
{
    if (captureEvent.WorkbookId == tracked.Id && captureEvent.Kind is CaptureEventKind.Captured or CaptureEventKind.NoTrackedChanges or CaptureEventKind.Failed)
        completion.TrySetResult(captureEvent);
};
await coordinator.StartAsync();
Directory.CreateDirectory(Path.GetDirectoryName(readySignal)!);
await File.WriteAllTextAsync(readySignal, tracked.Id.ToString("D"));
Console.WriteLine($"READY|{tracked.Id:D}");

var result = await completion.Task.WaitAsync(TimeSpan.FromSeconds(60));
var versions = await store.GetVersionsAsync(tracked.Id);
var errors = await store.GetErrorsAsync(tracked.Id);
var payload = new
{
    kind = result.Kind.ToString(),
    result.Message,
    result.ReportPath,
    versionCount = versions.Count,
    errorCount = errors.Count,
    cellChanges = result.Diff?.CellChanges.Count,
    sheetChanges = result.Diff?.SheetChanges.Count
};
await File.WriteAllTextAsync(resultPath, JsonSerializer.Serialize(payload, new JsonSerializerOptions { WriteIndented = true }));
Console.WriteLine($"RESULT|{result.Kind}|{result.ReportPath}");
return result.Kind == CaptureEventKind.Failed ? 1 : 0;
