using System.Collections.Concurrent;
using ExcelDiffTracker.Core;
using ExcelDiffTracker.Reporting;
using ExcelDiffTracker.Storage;

namespace ExcelDiffTracker.Tracking;

public sealed class TrackingCoordinator : IAsyncDisposable
{
    private static readonly TimeSpan DefaultReconciliationInterval = TimeSpan.FromSeconds(10);
    private static readonly TimeSpan EventSettleDelay = TimeSpan.FromMilliseconds(150);
    private readonly HistoryStore _store;
    private readonly WorkbookExtractor _extractor;
    private readonly WorkbookDiffer _differ;
    private readonly StableWorkbookCopy _stableCopy;
    private readonly MarkdownReportWriter _reportWriter;
    private readonly TimeSpan _reconciliationInterval;
    private readonly ConcurrentDictionary<Guid, SemaphoreSlim> _workbookGates = new();
    private readonly ConcurrentDictionary<Guid, WorkerState> _workers = new();
    private readonly ConcurrentDictionary<Guid, FileSystemWatcher> _watchers = new();
    private readonly ConcurrentDictionary<Guid, byte> _suspended = new();
    private readonly ConcurrentDictionary<Guid, byte> _purging = new();
    private readonly CancellationTokenSource _lifetime = new();
    private PeriodicTimer? _timer;
    private Task? _reconciliationTask;
    private bool _started;

    public TrackingCoordinator(
        HistoryStore store,
        WorkbookExtractor? extractor = null,
        WorkbookDiffer? differ = null,
        StableWorkbookCopy? stableCopy = null,
        MarkdownReportWriter? reportWriter = null,
        TimeSpan? reconciliationInterval = null)
    {
        _store = store ?? throw new ArgumentNullException(nameof(store));
        _extractor = extractor ?? new WorkbookExtractor();
        _differ = differ ?? new WorkbookDiffer();
        _stableCopy = stableCopy ?? new StableWorkbookCopy();
        _reportWriter = reportWriter ?? new MarkdownReportWriter();
        _reconciliationInterval = reconciliationInterval ?? DefaultReconciliationInterval;
    }

    public event EventHandler<CaptureEvent>? CaptureOccurred;

    public Task InitializeAsync(CancellationToken cancellationToken = default) => _store.InitializeAsync(cancellationToken);

    public async Task<TrackedWorkbook> AddWorkbookAsync(string workbookPath, string reportDirectory, CancellationToken cancellationToken = default)
    {
        ValidateSupportedPath(workbookPath);
        Directory.CreateDirectory(Path.GetFullPath(reportDirectory));
        using var stable = await _stableCopy.CreateAsync(workbookPath, cancellationToken).ConfigureAwait(false);
        var snapshot = await Task.Run(() => _extractor.Extract(stable.TemporaryPath), cancellationToken).ConfigureAwait(false);
        var workbook = await _store.AddBaselineAsync(
            Path.GetFullPath(workbookPath),
            Path.GetFullPath(reportDirectory),
            stable.Sha256,
            stable.SourceLastWriteUtc,
            snapshot,
            cancellationToken).ConfigureAwait(false);

        if (_started)
        {
            EnsureWorker(workbook.Id);
            AddWatcher(workbook);
        }
        Raise(new CaptureEvent
        {
            WorkbookId = workbook.Id,
            WorkbookPath = workbook.Path,
            Kind = CaptureEventKind.BaselineCreated,
            OccurredUtc = DateTime.UtcNow,
            Message = "Tracking started. The first snapshot was saved silently as the baseline."
        });
        return workbook;
    }

    public async Task StartAsync(CancellationToken cancellationToken = default)
    {
        if (_started)
            return;
        _started = true;
        _timer = new PeriodicTimer(_reconciliationInterval);
        _reconciliationTask = ReconcileLoopAsync(_lifetime.Token);
        await DispatchReconciliationAsync(cancellationToken).ConfigureAwait(false);
    }

    public async Task SetEnabledAsync(Guid workbookId, bool enabled, CancellationToken cancellationToken = default)
    {
        var gate = _workbookGates.GetOrAdd(workbookId, _ => new SemaphoreSlim(1, 1));
        if (!enabled)
        {
            _suspended[workbookId] = 0;
            RemoveWatcher(workbookId);
        }

        await gate.WaitAsync(cancellationToken).ConfigureAwait(false);
        TrackedWorkbook? workbook;
        try
        {
            await _store.SetEnabledAsync(workbookId, enabled, cancellationToken).ConfigureAwait(false);
            workbook = await _store.GetTrackedWorkbookAsync(workbookId, cancellationToken).ConfigureAwait(false);
        }
        finally
        {
            gate.Release();
        }
        if (workbook is null)
            return;

        if (enabled)
        {
            _suspended.TryRemove(workbookId, out _);
            EnsureWorker(workbookId);
            AddWatcher(workbook);
            ScheduleCapture(workbookId);
        }
        Raise(new CaptureEvent
        {
            WorkbookId = workbook.Id,
            WorkbookPath = workbook.Path,
            Kind = CaptureEventKind.StatusChanged,
            OccurredUtc = DateTime.UtcNow,
            Message = enabled ? "Tracking resumed." : "Tracking paused. History was retained."
        });
    }

    public async Task PurgeAsync(Guid workbookId, bool confirmed, CancellationToken cancellationToken = default)
    {
        if (!confirmed)
            throw new InvalidOperationException("Permanent deletion requires explicit confirmation.");
        _purging[workbookId] = 0;
        _suspended[workbookId] = 0;
        RemoveWatcher(workbookId);
        await RemoveWorkerAsync(workbookId).ConfigureAwait(false);
        var gate = _workbookGates.GetOrAdd(workbookId, _ => new SemaphoreSlim(1, 1));
        await gate.WaitAsync(cancellationToken).ConfigureAwait(false);
        try
        {
            await _store.MarkReportsPendingAsync(workbookId, cancellationToken).ConfigureAwait(false);
            var reports = await _store.GetReportPathsAsync(workbookId, cancellationToken).ConfigureAwait(false);
            foreach (var report in reports)
                DeleteReportOrThrow(report);
            _ = await _store.PurgeAsync(workbookId, confirmed: true, cancellationToken).ConfigureAwait(false);
        }
        finally
        {
            gate.Release();
            _purging.TryRemove(workbookId, out _);
            _suspended.TryRemove(workbookId, out _);
            _workbookGates.TryRemove(workbookId, out _);
        }
    }

    public async Task ReconcileAllAsync(CancellationToken cancellationToken = default)
    {
        var workbooks = await _store.GetTrackedWorkbooksAsync(cancellationToken).ConfigureAwait(false);
        await Task.WhenAll(workbooks.Where(item => item.IsEnabled).Select(item => CaptureIfChangedAsync(item.Id, cancellationToken))).ConfigureAwait(false);
    }

    public async Task CaptureIfChangedAsync(Guid workbookId, CancellationToken cancellationToken = default)
    {
        if (_purging.ContainsKey(workbookId) || _suspended.ContainsKey(workbookId))
            return;
        var gate = _workbookGates.GetOrAdd(workbookId, _ => new SemaphoreSlim(1, 1));
        await gate.WaitAsync(cancellationToken).ConfigureAwait(false);
        string? attemptedSha256 = null;
        var stage = "Opening workbook";
        try
        {
            if (_purging.ContainsKey(workbookId) || _suspended.ContainsKey(workbookId))
                return;
            var workbook = await _store.GetTrackedWorkbookAsync(workbookId, cancellationToken).ConfigureAwait(false);
            if (workbook is null || !workbook.IsEnabled)
                return;

            stage = "Recovering Markdown report";
            await RecoverPendingReportsAsync(workbook.Id, cancellationToken).ConfigureAwait(false);
            await _store.SetProcessingStatusAsync(workbook.Id, TrackingStatus.Processing, cancellationToken: cancellationToken).ConfigureAwait(false);
            Raise(new CaptureEvent
            {
                WorkbookId = workbook.Id,
                WorkbookPath = workbook.Path,
                Kind = CaptureEventKind.StatusChanged,
                OccurredUtc = DateTime.UtcNow,
                Message = "Processing workbook save…"
            });

            stage = "Waiting for a stable save";
            using var stable = await _stableCopy.CreateAsync(workbook.Path, cancellationToken).ConfigureAwait(false);
            attemptedSha256 = stable.Sha256;
            if (string.Equals(stable.Sha256, workbook.CurrentHash, StringComparison.OrdinalIgnoreCase))
            {
                await _store.SetProcessingStatusAsync(workbook.Id, TrackingStatus.Active, cancellationToken: cancellationToken).ConfigureAwait(false);
                Raise(new CaptureEvent
                {
                    WorkbookId = workbook.Id,
                    WorkbookPath = workbook.Path,
                    Kind = CaptureEventKind.DuplicateIgnored,
                    OccurredUtc = DateTime.UtcNow,
                    Message = "Workbook is unchanged; duplicate file event ignored."
                });
                return;
            }

            stage = "Reading workbook content";
            var current = await Task.Run(() => _extractor.Extract(stable.TemporaryPath), cancellationToken).ConfigureAwait(false);
            var previous = await _store.LoadCurrentSnapshotAsync(workbook, cancellationToken).ConfigureAwait(false);
            var diff = await Task.Run(() => _differ.Compare(previous, current), cancellationToken).ConfigureAwait(false);
            var sequence = workbook.CurrentSequence + 1;
            var summary = BuildSummary(diff);
            var reportPath = BuildReportPath(workbook, sequence, stable.SourceLastWriteUtc, stable.Sha256);

            stage = "Committing local history";
            var version = await _store.CommitVersionAsync(
                workbook,
                current,
                diff,
                stable.Sha256,
                stable.SourceLastWriteUtc,
                reportPath,
                summary,
                cancellationToken).ConfigureAwait(false);

            stage = "Writing Markdown report";
            var reportResult = await WriteReportAsync(version, diff, includeAll: false, cancellationToken).ConfigureAwait(false);
            await _store.MarkReportReadyAsync(version.Id, cancellationToken).ConfigureAwait(false);
            Raise(new CaptureEvent
            {
                WorkbookId = workbook.Id,
                WorkbookPath = workbook.Path,
                Kind = diff.HasTrackedChanges ? CaptureEventKind.Captured : CaptureEventKind.NoTrackedChanges,
                OccurredUtc = version.CapturedUtc,
                Message = reportResult.WasTruncated ? $"{summary} Automatic report truncated; full history retained." : summary,
                Diff = diff,
                ReportPath = reportResult.Path
            });
        }
        catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested || _lifetime.IsCancellationRequested)
        {
            throw;
        }
        catch (Exception exception)
        {
            var workbook = await SafeGetWorkbookAsync(workbookId).ConfigureAwait(false);
            if (workbook is not null)
            {
                var status = File.Exists(workbook.Path) ? TrackingStatus.Warning : TrackingStatus.Missing;
                var category = ClassifyError(exception, status);
                var fingerprint = $"{attemptedSha256 ?? "no-hash"}:{stage}:{category}:{exception.HResult}";
                try
                {
                    await _store.RecordErrorAsync(workbookId, fingerprint, attemptedSha256, $"{category} — {stage}", exception.Message, status, CancellationToken.None).ConfigureAwait(false);
                }
                catch
                {
                    // A secondary database error must not terminate the polling worker.
                }
                Raise(new CaptureEvent
                {
                    WorkbookId = workbookId,
                    WorkbookPath = workbook.Path,
                    Kind = CaptureEventKind.Failed,
                    OccurredUtc = DateTime.UtcNow,
                    Message = $"{stage}: {exception.Message}",
                    Exception = exception
                });
            }
        }
        finally
        {
            gate.Release();
        }
    }

    public async Task<ReportWriteResult> ExportFullReportAsync(VersionRecord version, string destinationPath, CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(version);
        _ = await _store.GetTrackedWorkbookAsync(version.WorkbookId, cancellationToken).ConfigureAwait(false)
            ?? throw new InvalidOperationException("The tracked workbook no longer exists.");
        var diff = await _store.GetVersionDiffAsync(version.Id, cancellationToken).ConfigureAwait(false);
        return await _reportWriter.WriteAsync(
            destinationPath,
            ContextFor(version),
            diff,
            includeAll: true,
            cancellationToken: cancellationToken).ConfigureAwait(false);
    }

    public async ValueTask DisposeAsync()
    {
        if (_lifetime.IsCancellationRequested)
            return;
        _lifetime.Cancel();
        _timer?.Dispose();
        foreach (var watcher in _watchers.Values)
            watcher.Dispose();
        _watchers.Clear();
        foreach (var worker in _workers.Values)
            worker.Cancel();
        if (_reconciliationTask is not null)
        {
            try { await _reconciliationTask.ConfigureAwait(false); }
            catch (OperationCanceledException) { }
        }
        var tasks = _workers.Values.Select(worker => worker.Task).Where(task => task is not null).Cast<Task>().ToArray();
        try { await Task.WhenAll(tasks).ConfigureAwait(false); }
        catch (OperationCanceledException) { }
        foreach (var worker in _workers.Values)
            worker.Dispose();
        foreach (var gate in _workbookGates.Values)
            gate.Dispose();
        _lifetime.Dispose();
    }

    private WorkerState EnsureWorker(Guid workbookId)
    {
        return _workers.GetOrAdd(workbookId, id =>
        {
            var worker = new WorkerState(_lifetime.Token);
            worker.Task = WorkerLoopAsync(id, worker);
            return worker;
        });
    }

    private async Task WorkerLoopAsync(Guid workbookId, WorkerState worker)
    {
        try
        {
            while (true)
            {
                await worker.Signal.WaitAsync(worker.Token).ConfigureAwait(false);
                await Task.Delay(EventSettleDelay, worker.Token).ConfigureAwait(false);
                do
                {
                    Interlocked.Exchange(ref worker.Dirty, 0);
                    await CaptureIfChangedAsync(workbookId, worker.Token).ConfigureAwait(false);
                }
                while (Volatile.Read(ref worker.Dirty) != 0);
            }
        }
        catch (OperationCanceledException) when (worker.Token.IsCancellationRequested)
        {
        }
        catch (Exception exception)
        {
            var workbook = await SafeGetWorkbookAsync(workbookId).ConfigureAwait(false);
            if (workbook is not null)
            {
                Raise(new CaptureEvent
                {
                    WorkbookId = workbookId,
                    WorkbookPath = workbook.Path,
                    Kind = CaptureEventKind.Failed,
                    OccurredUtc = DateTime.UtcNow,
                    Message = $"Tracking worker restarted after an internal error: {exception.Message}",
                    Exception = exception
                });
            }
            if (!_lifetime.IsCancellationRequested && !_purging.ContainsKey(workbookId))
            {
                _workers.TryRemove(workbookId, out _);
                ScheduleCapture(workbookId);
            }
        }
    }

    private async Task RemoveWorkerAsync(Guid workbookId)
    {
        if (!_workers.TryRemove(workbookId, out var worker))
            return;
        worker.Cancel();
        if (worker.Task is not null)
        {
            try { await worker.Task.ConfigureAwait(false); }
            catch (OperationCanceledException) { }
        }
        worker.Dispose();
    }

    private async Task ReconcileLoopAsync(CancellationToken cancellationToken)
    {
        while (!cancellationToken.IsCancellationRequested)
        {
            try
            {
                if (_timer is null || !await _timer.WaitForNextTickAsync(cancellationToken).ConfigureAwait(false))
                    return;
                await DispatchReconciliationAsync(cancellationToken).ConfigureAwait(false);
            }
            catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested)
            {
                return;
            }
            catch
            {
                await Task.Delay(TimeSpan.FromSeconds(1), cancellationToken).ConfigureAwait(false);
            }
        }
    }

    private async Task DispatchReconciliationAsync(CancellationToken cancellationToken)
    {
        var workbooks = await _store.GetTrackedWorkbooksAsync(cancellationToken).ConfigureAwait(false);
        foreach (var workbook in workbooks.Where(item => item.IsEnabled))
        {
            EnsureWorker(workbook.Id);
            AddWatcher(workbook);
            ScheduleCapture(workbook.Id);
        }
    }

    private void ScheduleCapture(Guid workbookId)
    {
        if (_lifetime.IsCancellationRequested || _purging.ContainsKey(workbookId) || _suspended.ContainsKey(workbookId))
            return;
        var worker = EnsureWorker(workbookId);
        Interlocked.Exchange(ref worker.Dirty, 1);
        if (worker.Signal.CurrentCount == 0)
        {
            try { worker.Signal.Release(); }
            catch (SemaphoreFullException) { }
        }
    }

    private void AddWatcher(TrackedWorkbook workbook)
    {
        if (_watchers.ContainsKey(workbook.Id) || _suspended.ContainsKey(workbook.Id) || _purging.ContainsKey(workbook.Id))
            return;
        var directory = Path.GetDirectoryName(workbook.Path);
        if (string.IsNullOrWhiteSpace(directory) || !Directory.Exists(directory))
            return;

        var watcher = new FileSystemWatcher(directory, Path.GetFileName(workbook.Path))
        {
            IncludeSubdirectories = false,
            NotifyFilter = NotifyFilters.FileName | NotifyFilters.LastWrite | NotifyFilters.Size | NotifyFilters.CreationTime,
            InternalBufferSize = 32 * 1024
        };
        watcher.Changed += (_, _) => ScheduleCapture(workbook.Id);
        watcher.Created += (_, _) => ScheduleCapture(workbook.Id);
        watcher.Renamed += (_, _) => ScheduleCapture(workbook.Id);
        watcher.Deleted += (_, _) => ScheduleCapture(workbook.Id);
        watcher.Error += (_, _) =>
        {
            RemoveWatcher(workbook.Id);
            ScheduleCapture(workbook.Id);
        };
        if (_watchers.TryAdd(workbook.Id, watcher))
            watcher.EnableRaisingEvents = true;
        else
            watcher.Dispose();
    }

    private void RemoveWatcher(Guid workbookId)
    {
        if (_watchers.TryRemove(workbookId, out var watcher))
        {
            watcher.EnableRaisingEvents = false;
            watcher.Dispose();
        }
    }

    private async Task RecoverPendingReportsAsync(Guid workbookId, CancellationToken cancellationToken)
    {
        foreach (var version in await _store.GetPendingReportsAsync(workbookId, cancellationToken).ConfigureAwait(false))
        {
            var diff = await _store.GetVersionDiffAsync(version.Id, cancellationToken).ConfigureAwait(false);
            _ = await WriteReportAsync(version, diff, includeAll: false, cancellationToken).ConfigureAwait(false);
            await _store.MarkReportReadyAsync(version.Id, cancellationToken).ConfigureAwait(false);
        }
    }

    private Task<ReportWriteResult> WriteReportAsync(VersionRecord version, WorkbookDiff diff, bool includeAll, CancellationToken cancellationToken) =>
        _reportWriter.WriteAsync(
            version.ReportPath ?? throw new InvalidDataException("Version has no Markdown report path."),
            ContextFor(version),
            diff,
            includeAll: includeAll,
            cancellationToken: cancellationToken);

    private static ReportContext ContextFor(VersionRecord version) => new()
    {
        WorkbookPath = version.WorkbookPath,
        Sequence = version.Sequence,
        CapturedUtc = version.CapturedUtc,
        FileLastWriteUtc = version.FileLastWriteUtc,
        PreviousSha256 = version.PreviousSha256,
        CurrentSha256 = version.Sha256
    };

    private async Task<TrackedWorkbook?> SafeGetWorkbookAsync(Guid id)
    {
        try { return await _store.GetTrackedWorkbookAsync(id, CancellationToken.None).ConfigureAwait(false); }
        catch { return null; }
    }

    private void Raise(CaptureEvent captureEvent)
    {
        var handlers = CaptureOccurred;
        if (handlers is null)
            return;
        foreach (EventHandler<CaptureEvent> handler in handlers.GetInvocationList())
        {
            try { handler(this, captureEvent); }
            catch { }
        }
    }

    private static string BuildSummary(WorkbookDiff diff) => !diff.HasTrackedChanges
        ? "No tracked changes"
        : $"{diff.CellChanges.Count:N0} cells, {diff.SheetChanges.Count:N0} sheet changes";

    private static string BuildReportPath(TrackedWorkbook workbook, long sequence, DateTime savedUtc, string sha256)
    {
        var workbookName = SanitizeFileName(Path.GetFileNameWithoutExtension(workbook.Path));
        var folder = Path.Combine(workbook.ReportDirectory, $"{workbookName}-{workbook.Id.ToString("N")[..8]}");
        return Path.Combine(folder, $"{sequence:D6}_{savedUtc.ToLocalTime():yyyy-MM-dd_HHmmss}_{sha256[..12]}.md");
    }

    private static string SanitizeFileName(string value)
    {
        var invalid = Path.GetInvalidFileNameChars();
        var result = new string(value.Select(character => invalid.Contains(character) ? '_' : character).ToArray()).Trim();
        return string.IsNullOrWhiteSpace(result) ? "Workbook" : result;
    }

    private static void ValidateSupportedPath(string path)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(path);
        var extension = Path.GetExtension(path);
        if (!extension.Equals(".xlsx", StringComparison.OrdinalIgnoreCase) && !extension.Equals(".xlsm", StringComparison.OrdinalIgnoreCase))
            throw new UnsupportedWorkbookException("Choose an .xlsx or .xlsm workbook. Legacy .xls and binary .xlsb files are not supported.");
    }

    private static string ClassifyError(Exception exception, TrackingStatus status)
    {
        if (status == TrackingStatus.Missing)
            return "Missing workbook";
        return exception switch
        {
            UnsupportedWorkbookException => "Unsupported workbook",
            UnsafeWorkbookException => "Corrupt, encrypted, or unsafe workbook",
            UnauthorizedAccessException => "Permission denied",
            IOException => "Workbook temporarily unavailable",
            _ => "Capture failed"
        };
    }

    private static void DeleteReportOrThrow(string path)
    {
        if (!File.Exists(path))
            return;
        File.Delete(path);
    }

    private sealed class WorkerState : IDisposable
    {
        private readonly CancellationTokenSource _cancellation;

        public WorkerState(CancellationToken lifetime)
        {
            _cancellation = CancellationTokenSource.CreateLinkedTokenSource(lifetime);
        }

        public SemaphoreSlim Signal { get; } = new(0, 1);
        public CancellationToken Token => _cancellation.Token;
        public Task? Task { get; set; }
        public int Dirty;
        public void Cancel() => _cancellation.Cancel();
        public void Dispose()
        {
            _cancellation.Dispose();
            Signal.Dispose();
        }
    }
}
