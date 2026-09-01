using System.Security.Cryptography;
using DocumentFormat.OpenXml.Spreadsheet;
using ExcelDiffTracker.Core;
using ExcelDiffTracker.Reporting;
using ExcelDiffTracker.Storage;
using ExcelDiffTracker.Tracking;
using Microsoft.Data.Sqlite;

namespace ExcelDiffTracker.Tests;

public sealed class StorageAndTrackingTests
{
    [Fact]
    public async Task StableCopyReadsWorkbookHeldOpenByACompatibleWriter()
    {
        using var directory = new TestDirectory();
        var workbookPath = Path.Combine(directory.Path, "open-in-excel.xlsx");
        WorkbookFixture.Create(workbookPath, false, new FixtureSheet(1, "Sheet1", null, new FixtureCell("A1", "42")));
        var expectedHash = Hash(workbookPath);
        var expectedLength = new FileInfo(workbookPath).Length;

        await using var writer = new FileStream(
            workbookPath,
            FileMode.Open,
            FileAccess.ReadWrite,
            FileShare.ReadWrite | FileShare.Delete);
        using var stable = await new StableWorkbookCopy(
            TimeSpan.FromMilliseconds(10),
            TimeSpan.FromMilliseconds(10),
            TimeSpan.FromSeconds(1)).CreateAsync(workbookPath);

        Assert.Equal(expectedHash, stable.Sha256);
        Assert.Equal(expectedLength, stable.SourceLength);
        Assert.Equal("42", new WorkbookExtractor().Extract(stable.TemporaryPath).Sheets[1].Cells["A1"].LiteralValue);
    }

    [Fact]
    public async Task BaselineIsSilentAndFirstSaveIsCommittedTransactionally()
    {
        using var directory = new TestDirectory();
        var database = Path.Combine(directory.Path, "history.db");
        var workbookPath = Path.Combine(directory.Path, "book.xlsx");
        var reportDirectory = Path.Combine(directory.Path, "reports");
        WorkbookFixture.Create(workbookPath, false, new FixtureSheet(1, "Sheet1", null, new FixtureCell("A1", "1")));
        var extractor = new WorkbookExtractor();
        var baselineSnapshot = extractor.Extract(workbookPath);
        var store = new HistoryStore(database);
        await store.InitializeAsync();
        var baseline = await store.AddBaselineAsync(workbookPath, reportDirectory, Hash(workbookPath), File.GetLastWriteTimeUtc(workbookPath), baselineSnapshot);

        Assert.Equal(0, baseline.CurrentSequence);
        Assert.Empty(await store.GetVersionsAsync());

        WorkbookFixture.Create(workbookPath, false, new FixtureSheet(1, "Sheet1", null, new FixtureCell("A1", "2")));
        var current = extractor.Extract(workbookPath);
        var diff = new WorkbookDiffer().Compare(baselineSnapshot, current);
        var reportPath = Path.Combine(reportDirectory, "version.md");
        Directory.CreateDirectory(reportDirectory);
        await File.WriteAllTextAsync(reportPath, "report");
        var version = await store.CommitVersionAsync(baseline, current, diff, Hash(workbookPath), File.GetLastWriteTimeUtc(workbookPath), reportPath, "1 cell");

        Assert.Equal(1, version.Sequence);
        Assert.Equal(baseline.CurrentHash, version.PreviousSha256);
        Assert.Single(await store.GetVersionsAsync());
        Assert.Equal("2", (await store.LoadCurrentSnapshotAsync((await store.GetTrackedWorkbookAsync(baseline.Id))!)).Sheets[1].Cells["A1"].LiteralValue);
    }

    [Fact]
    public async Task HybridCaptureWritesStyleOnlyReportDeduplicatesAndKeepsLastGoodBaselineOnCorruption()
    {
        using var directory = new TestDirectory();
        var workbookPath = Path.Combine(directory.Path, "book.xlsx");
        var reports = Path.Combine(directory.Path, "reports");
        var store = new HistoryStore(Path.Combine(directory.Path, "history.db"));
        var coordinator = new TrackingCoordinator(
            store,
            stableCopy: new StableWorkbookCopy(TimeSpan.FromMilliseconds(10), TimeSpan.FromMilliseconds(10), TimeSpan.FromMilliseconds(250)),
            reconciliationInterval: TimeSpan.FromHours(1));
        await coordinator.InitializeAsync();
        WorkbookFixture.Create(workbookPath, false, new FixtureSheet(1, "Sheet1", null, new FixtureCell("A1", "42", StyleIndex: 0)));
        var tracked = await coordinator.AddWorkbookAsync(workbookPath, reports);

        WorkbookFixture.Create(workbookPath, false, new FixtureSheet(1, "Sheet1", null, new FixtureCell("A1", "42", StyleIndex: 1)));
        await coordinator.CaptureIfChangedAsync(tracked.Id);
        var versions = await store.GetVersionsAsync(tracked.Id);
        var styleVersion = Assert.Single(versions);
        Assert.Equal(0, styleVersion.CellChangeCount);
        Assert.Contains("No tracked changes", await File.ReadAllTextAsync(styleVersion.ReportPath!));

        await coordinator.CaptureIfChangedAsync(tracked.Id);
        Assert.Single(await store.GetVersionsAsync(tracked.Id));

        var lastGood = await store.GetTrackedWorkbookAsync(tracked.Id);
        await File.WriteAllTextAsync(workbookPath, "not an Open XML package");
        await coordinator.CaptureIfChangedAsync(tracked.Id);
        var afterFailure = await store.GetTrackedWorkbookAsync(tracked.Id);
        Assert.Equal(lastGood!.CurrentHash, afterFailure!.CurrentHash);
        Assert.Equal(lastGood.CurrentSequence, afterFailure.CurrentSequence);
        Assert.Single(await store.GetErrorsAsync(tracked.Id));
        await coordinator.DisposeAsync();
    }

    [Fact]
    public async Task NewerDatabaseSchemaIsRejectedBeforeMigration()
    {
        using var directory = new TestDirectory();
        var database = Path.Combine(directory.Path, "future.db");
        await using (var connection = new SqliteConnection($"Data Source={database}"))
        {
            await connection.OpenAsync();
            await using var command = connection.CreateCommand();
            command.CommandText = "PRAGMA user_version=99; CREATE TABLE future_data(value TEXT); INSERT INTO future_data(value) VALUES('keep');";
            await command.ExecuteNonQueryAsync();
        }

        var store = new HistoryStore(database);
        await Assert.ThrowsAsync<InvalidOperationException>(() => store.InitializeAsync());
        await using var verify = new SqliteConnection($"Data Source={database}");
        await verify.OpenAsync();
        await using var verifyCommand = verify.CreateCommand();
        verifyCommand.CommandText = "SELECT value FROM future_data;";
        Assert.Equal("keep", await verifyCommand.ExecuteScalarAsync());
    }

    [Fact]
    public async Task PendingReportIsRecoveredFromCanonicalHistory()
    {
        using var directory = new TestDirectory();
        var workbookPath = Path.Combine(directory.Path, "book.xlsx");
        var reportPath = Path.Combine(directory.Path, "reports", "pending.md");
        WorkbookFixture.Create(workbookPath, false, new FixtureSheet(1, "Sheet1", null, new FixtureCell("A1", "1")));
        var extractor = new WorkbookExtractor();
        var baselineSnapshot = extractor.Extract(workbookPath);
        var store = new HistoryStore(Path.Combine(directory.Path, "history.db"));
        await store.InitializeAsync();
        var baseline = await store.AddBaselineAsync(workbookPath, directory.Path, Hash(workbookPath), File.GetLastWriteTimeUtc(workbookPath), baselineSnapshot);
        WorkbookFixture.Create(workbookPath, false, new FixtureSheet(1, "Sheet1", null, new FixtureCell("A1", "2")));
        var current = extractor.Extract(workbookPath);
        var diff = new WorkbookDiffer().Compare(baselineSnapshot, current);
        var version = await store.CommitVersionAsync(baseline, current, diff, Hash(workbookPath), File.GetLastWriteTimeUtc(workbookPath), reportPath, "1 cell");
        Assert.Equal(ReportStatus.Pending, version.ReportStatus);
        Assert.False(File.Exists(reportPath));

        await using var coordinator = new TrackingCoordinator(store, stableCopy: new StableWorkbookCopy(TimeSpan.FromMilliseconds(10), TimeSpan.FromMilliseconds(10), TimeSpan.FromSeconds(1)));
        await coordinator.CaptureIfChangedAsync(baseline.Id);

        Assert.True(File.Exists(reportPath));
        Assert.Empty(await store.GetPendingReportsAsync(baseline.Id));
    }

    [Fact]
    public async Task WorkerCapturesRapidSequentialSavesWithoutCancelingInFlightCapture()
    {
        using var directory = new TestDirectory();
        var workbookPath = Path.Combine(directory.Path, "rapid.xlsx");
        WorkbookFixture.Create(workbookPath, false, new FixtureSheet(1, "Sheet1", null, new FixtureCell("A1", "1")));
        var store = new HistoryStore(Path.Combine(directory.Path, "history.db"));
        await using var coordinator = new TrackingCoordinator(
            store,
            stableCopy: new StableWorkbookCopy(TimeSpan.FromMilliseconds(20), TimeSpan.FromMilliseconds(10), TimeSpan.FromSeconds(2)),
            reconciliationInterval: TimeSpan.FromMilliseconds(100));
        await coordinator.InitializeAsync();
        var tracked = await coordinator.AddWorkbookAsync(workbookPath, Path.Combine(directory.Path, "reports"));
        await coordinator.StartAsync();

        WorkbookFixture.Create(workbookPath, false, new FixtureSheet(1, "Sheet1", null, new FixtureCell("A1", "2")));
        await WaitForVersionCountAsync(store, tracked.Id, 1);
        WorkbookFixture.Create(workbookPath, false, new FixtureSheet(1, "Sheet1", null, new FixtureCell("A1", "3")));
        await WaitForVersionCountAsync(store, tracked.Id, 2);

        var versions = await store.GetVersionsAsync(tracked.Id);
        Assert.Equal(2, versions.Count);
        Assert.Equal(2, versions.Select(item => item.Sha256).Distinct(StringComparer.OrdinalIgnoreCase).Count());
    }

    [Fact]
    public async Task PurgeWaitsForAnActiveCaptureAndLeavesNoHistory()
    {
        using var directory = new TestDirectory();
        var workbookPath = Path.Combine(directory.Path, "purge.xlsx");
        WorkbookFixture.Create(workbookPath, false, new FixtureSheet(1, "Sheet1", null, new FixtureCell("A1", "1")));
        var store = new HistoryStore(Path.Combine(directory.Path, "history.db"));
        await using var coordinator = new TrackingCoordinator(
            store,
            stableCopy: new StableWorkbookCopy(TimeSpan.FromMilliseconds(150), TimeSpan.FromMilliseconds(20), TimeSpan.FromSeconds(2)));
        await coordinator.InitializeAsync();
        var tracked = await coordinator.AddWorkbookAsync(workbookPath, Path.Combine(directory.Path, "reports"));
        WorkbookFixture.Create(workbookPath, false, new FixtureSheet(1, "Sheet1", null, new FixtureCell("A1", "2")));

        var capture = coordinator.CaptureIfChangedAsync(tracked.Id);
        await Task.Delay(30);
        await coordinator.PurgeAsync(tracked.Id, confirmed: true);
        await capture;

        Assert.Null(await store.GetTrackedWorkbookAsync(tracked.Id));
        Assert.Empty(await store.GetVersionsAsync(tracked.Id));
    }

    [Fact]
    public async Task WorkbooksCaptureIndependentlyAcrossAtomicReplacementAndLockedRetry()
    {
        using var directory = new TestDirectory();
        var firstPath = Path.Combine(directory.Path, "first.xlsx");
        var secondPath = Path.Combine(directory.Path, "second.xlsx");
        WorkbookFixture.Create(firstPath, false, new FixtureSheet(1, "Sheet1", null, new FixtureCell("A1", "1")));
        WorkbookFixture.Create(secondPath, false, new FixtureSheet(1, "Sheet1", null, new FixtureCell("A1", "10")));
        var store = new HistoryStore(Path.Combine(directory.Path, "history.db"));
        await using var coordinator = new TrackingCoordinator(
            store,
            stableCopy: new StableWorkbookCopy(TimeSpan.FromMilliseconds(20), TimeSpan.FromMilliseconds(20), TimeSpan.FromSeconds(2)),
            reconciliationInterval: TimeSpan.FromHours(1));
        await coordinator.InitializeAsync();
        var first = await coordinator.AddWorkbookAsync(firstPath, Path.Combine(directory.Path, "first-reports"));
        var second = await coordinator.AddWorkbookAsync(secondPath, Path.Combine(directory.Path, "second-reports"));

        var replacement = Path.Combine(directory.Path, "replacement.xlsx");
        WorkbookFixture.Create(replacement, false, new FixtureSheet(1, "Sheet1", null, new FixtureCell("A1", "2")));
        File.Move(replacement, firstPath, overwrite: true);
        WorkbookFixture.Create(secondPath, false, new FixtureSheet(1, "Sheet1", null, new FixtureCell("A1", "20")));

        Task lockedCapture;
        using (File.Open(secondPath, FileMode.Open, FileAccess.Read, FileShare.None))
        {
            lockedCapture = coordinator.CaptureIfChangedAsync(second.Id);
            await coordinator.CaptureIfChangedAsync(first.Id);
            Assert.Single(await store.GetVersionsAsync(first.Id));
            Assert.Empty(await store.GetVersionsAsync(second.Id));
        }
        await lockedCapture;

        Assert.Single(await store.GetVersionsAsync(second.Id));
    }

    [Fact]
    public async Task ReconciliationRecoversAfterLockOutlastsTimeoutWithoutAFileEventAndPreservesBaseline()
    {
        using var directory = new TestDirectory();
        var workbookPath = Path.Combine(directory.Path, "recover.xlsx");
        WorkbookFixture.Create(workbookPath, false, new FixtureSheet(1, "Sheet1", null, new FixtureCell("A1", "1")));
        var store = new HistoryStore(Path.Combine(directory.Path, "history.db"));
        await using var coordinator = new TrackingCoordinator(
            store,
            stableCopy: new StableWorkbookCopy(
                TimeSpan.FromMilliseconds(10),
                TimeSpan.FromMilliseconds(10),
                TimeSpan.FromMilliseconds(100)),
            reconciliationInterval: TimeSpan.FromMilliseconds(40));
        await coordinator.InitializeAsync();
        var baseline = await coordinator.AddWorkbookAsync(workbookPath, Path.Combine(directory.Path, "reports"));
        WorkbookFixture.Create(workbookPath, false, new FixtureSheet(1, "Sheet1", null, new FixtureCell("A1", "2")));

        await using var exclusiveWriter = new FileStream(
            workbookPath,
            FileMode.Open,
            FileAccess.ReadWrite,
            FileShare.None);
        await coordinator.StartAsync();
        await WaitForErrorCountAsync(store, baseline.Id, 1);

        var afterTimeout = await store.GetTrackedWorkbookAsync(baseline.Id);
        Assert.NotNull(afterTimeout);
        Assert.Equal(baseline.CurrentHash, afterTimeout.CurrentHash);
        Assert.Equal(baseline.CurrentSequence, afterTimeout.CurrentSequence);
        Assert.Empty(await store.GetVersionsAsync(baseline.Id));

        await exclusiveWriter.DisposeAsync();
        await WaitForVersionCountAsync(store, baseline.Id, 1);

        var version = Assert.Single(await store.GetVersionsAsync(baseline.Id));
        Assert.Equal(baseline.CurrentHash, version.PreviousSha256);
        var recovered = await store.GetTrackedWorkbookAsync(baseline.Id);
        Assert.NotNull(recovered);
        Assert.Equal("2", (await store.LoadCurrentSnapshotAsync(recovered)).Sheets[1].Cells["A1"].LiteralValue);
    }

    [Fact]
    public async Task RestartReconciliationCapturesAChangeMadeWhileTrackerWasStopped()
    {
        using var directory = new TestDirectory();
        var workbookPath = Path.Combine(directory.Path, "restart.xlsx");
        var databasePath = Path.Combine(directory.Path, "history.db");
        WorkbookFixture.Create(workbookPath, false, new FixtureSheet(1, "Sheet1", null, new FixtureCell("A1", "1")));
        var store = new HistoryStore(databasePath);
        Guid workbookId;
        await using (var firstCoordinator = new TrackingCoordinator(store, reconciliationInterval: TimeSpan.FromHours(1)))
        {
            await firstCoordinator.InitializeAsync();
            workbookId = (await firstCoordinator.AddWorkbookAsync(workbookPath, Path.Combine(directory.Path, "reports"))).Id;
        }

        WorkbookFixture.Create(workbookPath, false, new FixtureSheet(1, "Sheet1", null, new FixtureCell("A1", "2")));
        var restartedStore = new HistoryStore(databasePath);
        await using var restartedCoordinator = new TrackingCoordinator(
            restartedStore,
            stableCopy: new StableWorkbookCopy(TimeSpan.FromMilliseconds(10), TimeSpan.FromMilliseconds(10), TimeSpan.FromSeconds(1)),
            reconciliationInterval: TimeSpan.FromHours(1));
        await restartedCoordinator.InitializeAsync();
        await restartedCoordinator.StartAsync();
        await WaitForVersionCountAsync(restartedStore, workbookId, 1);

        Assert.Single(await restartedStore.GetVersionsAsync(workbookId));
    }

    private static async Task WaitForVersionCountAsync(HistoryStore store, Guid workbookId, int expected)
    {
        var deadline = DateTime.UtcNow.AddSeconds(8);
        while (DateTime.UtcNow < deadline)
        {
            if ((await store.GetVersionsAsync(workbookId)).Count >= expected)
                return;
            await Task.Delay(50);
        }
        Assert.Fail($"Expected {expected} versions before timeout.");
    }

    private static async Task WaitForErrorCountAsync(HistoryStore store, Guid workbookId, int expected)
    {
        var deadline = DateTime.UtcNow.AddSeconds(8);
        while (DateTime.UtcNow < deadline)
        {
            if ((await store.GetErrorsAsync(workbookId)).Count >= expected)
                return;
            await Task.Delay(50);
        }
        Assert.Fail($"Expected {expected} errors before timeout.");
    }

    private static string Hash(string path) => Convert.ToHexStringLower(SHA256.HashData(File.ReadAllBytes(path)));
}
