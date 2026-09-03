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
    public async Task SelectedBaselineReconstructsHistoryWithoutRewritingChronologicalEvidence()
    {
        using var directory = new TestDirectory();
        var workbookPath = Path.Combine(directory.Path, "scenario.xlsx");
        var reports = Path.Combine(directory.Path, "reports");
        var extractor = new WorkbookExtractor();
        var differ = new WorkbookDiffer();
        var store = new HistoryStore(Path.Combine(directory.Path, "history.db"));
        await store.InitializeAsync();

        WorkbookFixture.Create(workbookPath, false,
            new FixtureSheet(1, "Sheet1", null,
                new FixtureCell("A1", "1"),
                new FixtureCell("B1", "2", "A1*2")));
        var baselineSnapshot = extractor.Extract(workbookPath);
        var tracked = await store.AddBaselineAsync(workbookPath, reports, Hash(workbookPath), File.GetLastWriteTimeUtc(workbookPath), baselineSnapshot);

        WorkbookFixture.Create(workbookPath, false,
            new FixtureSheet(1, "Sheet1", null,
                new FixtureCell("A1", "2"),
                new FixtureCell("B1", "4", "A1*2")),
            new FixtureSheet(2, "Added", null, new FixtureCell("C1", "x", Type: CellValues.String)));
        var firstSnapshot = extractor.Extract(workbookPath);
        var firstReport = Path.Combine(reports, "first.md");
        Directory.CreateDirectory(reports);
        await File.WriteAllTextAsync(firstReport, "immutable first report");
        var first = await store.CommitVersionAsync(
            tracked,
            firstSnapshot,
            differ.Compare(baselineSnapshot, firstSnapshot),
            Hash(workbookPath),
            File.GetLastWriteTimeUtc(workbookPath),
            firstReport,
            "first");

        tracked = (await store.GetTrackedWorkbookAsync(tracked.Id))!;
        WorkbookFixture.Create(workbookPath, false,
            new FixtureSheet(1, "Inputs", null,
                new FixtureCell("A1", "3"),
                new FixtureCell("B1", "9", "A1*3")));
        var secondSnapshot = extractor.Extract(workbookPath);
        var secondReport = Path.Combine(reports, "second.md");
        await File.WriteAllTextAsync(secondReport, "immutable second report");
        var second = await store.CommitVersionAsync(
            tracked,
            secondSnapshot,
            differ.Compare(firstSnapshot, secondSnapshot),
            Hash(workbookPath),
            File.GetLastWriteTimeUtc(workbookPath),
            secondReport,
            "second");

        tracked = (await store.GetTrackedWorkbookAsync(tracked.Id))!;
        var scan0 = await store.LoadSnapshotAtSequenceAsync(tracked, 0);
        var scan1 = await store.LoadSnapshotAtSequenceAsync(tracked, 1);
        var scan2 = await store.LoadSnapshotAtSequenceAsync(tracked, 2);
        Assert.Equal("1", scan0.Snapshot.Sheets[1].Cells["A1"].LiteralValue);
        Assert.Equal("Sheet1", scan0.Snapshot.Sheets[1].Name);
        Assert.False(scan0.Snapshot.Sheets.ContainsKey(2));
        Assert.Equal("2", scan1.Snapshot.Sheets[1].Cells["A1"].LiteralValue);
        Assert.Equal("A1*2", scan1.Snapshot.Sheets[1].Cells["B1"].FormulaText);
        Assert.Equal("x", scan1.Snapshot.Sheets[2].Cells["C1"].LiteralValue);
        Assert.Equal("Inputs", scan2.Snapshot.Sheets[1].Name);
        Assert.Equal("A1*3", scan2.Snapshot.Sheets[1].Cells["B1"].FormulaText);

        var firstBytes = await File.ReadAllBytesAsync(firstReport);
        var secondBytes = await File.ReadAllBytesAsync(secondReport);
        await using var coordinator = new TrackingCoordinator(store);
        await coordinator.SetComparisonBaselineAsync(tracked.Id, ComparisonBaselineMode.FirstScan);
        var fromOriginal = await coordinator.GetSelectedComparisonAsync(tracked.Id);
        Assert.Equal(0, fromOriginal.Baseline.Sequence);
        Assert.Equal(2, fromOriginal.Current.Sequence);
        Assert.Contains(fromOriginal.Diff.SheetChanges, change => change.Kind == SheetChangeKind.Renamed);

        await coordinator.SetComparisonBaselineAsync(tracked.Id, ComparisonBaselineMode.SpecificScan, first.Id);
        var fromFirstSave = await coordinator.GetSelectedComparisonAsync(tracked.Id);
        Assert.Equal(1, fromFirstSave.Baseline.Sequence);
        Assert.Contains(fromFirstSave.Diff.SheetChanges, change => change.Kind == SheetChangeKind.Removed);
        Assert.Contains(fromFirstSave.Diff.CellChanges, change => change.Address == "B1" && change.Kinds.Contains(CellChangeKind.FormulaChanged));

        await coordinator.SetComparisonBaselineAsync(tracked.Id, ComparisonBaselineMode.Previous);
        Assert.Equal(1, (await coordinator.GetSelectedComparisonAsync(tracked.Id)).Baseline.Sequence);
        var derived = Path.Combine(reports, "derived.md");
        await coordinator.ExportSelectedComparisonAsync(tracked.Id, derived);
        var markdown = await File.ReadAllTextAsync(derived);
        Assert.Contains("derived comparison view", markdown);
        Assert.Contains("Baseline mode: Previous", markdown);
        Assert.Contains($"Baseline SHA-256: `{second.PreviousSha256}`", markdown);
        Assert.Equal(firstBytes, await File.ReadAllBytesAsync(firstReport));
        Assert.Equal(secondBytes, await File.ReadAllBytesAsync(secondReport));

        var collision = await Assert.ThrowsAsync<InvalidOperationException>(() =>
            coordinator.ExportSelectedComparisonAsync(tracked.Id, firstReport));
        Assert.Contains("automatic chronological report", collision.Message);
        Assert.Equal(firstBytes, await File.ReadAllBytesAsync(firstReport));
        Assert.Equal(secondBytes, await File.ReadAllBytesAsync(secondReport));
    }

    [Fact]
    public async Task HistoricalReconstructionPreservesNoTrackedChangeScansAcrossRestart()
    {
        using var directory = new TestDirectory();
        var workbookPath = Path.Combine(directory.Path, "style-only.xlsx");
        var databasePath = Path.Combine(directory.Path, "history.db");
        var reports = Path.Combine(directory.Path, "reports");
        WorkbookFixture.Create(workbookPath, false, new FixtureSheet(1, "Sheet1", null, new FixtureCell("A1", "42", StyleIndex: 0)));
        var extractor = new WorkbookExtractor();
        var store = new HistoryStore(databasePath);
        await store.InitializeAsync();
        var baselineSnapshot = extractor.Extract(workbookPath);
        var tracked = await store.AddBaselineAsync(workbookPath, reports, Hash(workbookPath), File.GetLastWriteTimeUtc(workbookPath), baselineSnapshot);

        WorkbookFixture.Create(workbookPath, false, new FixtureSheet(1, "Sheet1", null, new FixtureCell("A1", "42", StyleIndex: 1)));
        var styleOnlySnapshot = extractor.Extract(workbookPath);
        var diff = new WorkbookDiffer().Compare(baselineSnapshot, styleOnlySnapshot);
        Assert.False(diff.HasTrackedChanges);
        var version = await store.CommitVersionAsync(
            tracked,
            styleOnlySnapshot,
            diff,
            Hash(workbookPath),
            File.GetLastWriteTimeUtc(workbookPath),
            Path.Combine(reports, "style-only.md"),
            "No tracked changes");
        await store.SetComparisonBaselineAsync(tracked.Id, ComparisonBaselineMode.SpecificScan, version.Id);

        var reopened = new HistoryStore(databasePath);
        await reopened.InitializeAsync();
        tracked = (await reopened.GetTrackedWorkbookAsync(tracked.Id))!;
        Assert.Equal(ComparisonBaselineMode.SpecificScan, tracked.ComparisonBaselineMode);
        Assert.Equal(version.Id, tracked.SpecificBaselineVersionId);
        Assert.Equal("42", (await reopened.LoadSnapshotAtSequenceAsync(tracked, 0)).Snapshot.Sheets[1].Cells["A1"].LiteralValue);
        Assert.Equal("42", (await reopened.LoadSnapshotAtSequenceAsync(tracked, 1)).Snapshot.Sheets[1].Cells["A1"].LiteralValue);
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
    public async Task WorkerCapturesSequentialCompatibleSavesWithoutDuplicates()
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
        var baselineHash = tracked.CurrentHash;
        var failedCaptureCount = 0;
        coordinator.CaptureOccurred += (_, capture) =>
        {
            if (capture.Kind == CaptureEventKind.Failed)
                Interlocked.Increment(ref failedCaptureCount);
        };
        await coordinator.StartAsync();

        WorkbookFixture.CreateWithCompatibleSharing(workbookPath, false, new FixtureSheet(1, "Sheet1", null, new FixtureCell("A1", "2")));
        var firstSaveHash = Hash(workbookPath);
        await WaitForReadyVersionCountAsync(store, tracked.Id, 1);
        WorkbookFixture.CreateWithCompatibleSharing(workbookPath, false, new FixtureSheet(1, "Sheet1", null, new FixtureCell("A1", "3")));
        var secondSaveHash = Hash(workbookPath);
        await WaitForReadyVersionCountAsync(store, tracked.Id, 2);
        await Task.Delay(600);

        var versions = (await store.GetVersionsAsync(tracked.Id)).OrderBy(item => item.Sequence).ToArray();
        Assert.Equal(2, versions.Length);
        Assert.Equal(2, versions.Select(item => item.Sha256).Distinct(StringComparer.OrdinalIgnoreCase).Count());
        Assert.Equal(new long[] { 1, 2 }, versions.Select(item => item.Sequence));
        Assert.Equal(baselineHash, versions[0].PreviousSha256);
        Assert.Equal(firstSaveHash, versions[0].Sha256);
        Assert.Equal(firstSaveHash, versions[1].PreviousSha256);
        Assert.Equal(secondSaveHash, versions[1].Sha256);
        Assert.All(versions, version => Assert.Equal(ReportStatus.Ready, version.ReportStatus));

        var firstDiff = await store.GetVersionDiffAsync(versions[0].Id);
        var firstDelta = Assert.Single(firstDiff.CellChanges);
        Assert.Empty(firstDiff.SheetChanges);
        Assert.Equal("A1", firstDelta.Address);
        Assert.Equal(new[] { CellChangeKind.LiteralChanged }, firstDelta.Kinds);
        Assert.Equal("1", firstDelta.Before?.LiteralValue);
        Assert.Equal("2", firstDelta.After?.LiteralValue);

        var secondDiff = await store.GetVersionDiffAsync(versions[1].Id);
        var secondDelta = Assert.Single(secondDiff.CellChanges);
        Assert.Empty(secondDiff.SheetChanges);
        Assert.Equal("A1", secondDelta.Address);
        Assert.Equal(new[] { CellChangeKind.LiteralChanged }, secondDelta.Kinds);
        Assert.Equal("2", secondDelta.Before?.LiteralValue);
        Assert.Equal("3", secondDelta.After?.LiteralValue);

        var finalWorkbook = await store.GetTrackedWorkbookAsync(tracked.Id);
        Assert.NotNull(finalWorkbook);
        Assert.Equal(2, finalWorkbook.CurrentSequence);
        Assert.Equal(secondSaveHash, finalWorkbook.CurrentHash);
        Assert.Equal("3", (await store.LoadCurrentSnapshotAsync(finalWorkbook)).Sheets[1].Cells["A1"].LiteralValue);
        Assert.Empty(await store.GetErrorsAsync(tracked.Id));
        Assert.Equal(0, Volatile.Read(ref failedCaptureCount));
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

    private static async Task WaitForReadyVersionCountAsync(HistoryStore store, Guid workbookId, int expected)
    {
        var deadline = DateTime.UtcNow.AddSeconds(8);
        while (DateTime.UtcNow < deadline)
        {
            var versions = await store.GetVersionsAsync(workbookId);
            if (versions.Count == expected && versions.All(version => version.ReportStatus == ReportStatus.Ready))
                return;
            await Task.Delay(50);
        }
        Assert.Fail($"Expected exactly {expected} ready versions before timeout.");
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
