using System.Collections.ObjectModel;
using System.Globalization;
using System.Text.Json;
using ExcelDiffTracker.Core;
using Microsoft.Data.Sqlite;

namespace ExcelDiffTracker.Storage;

public sealed class HistoryStore
{
    private const int CurrentSchemaVersion = 1;
    private const int PayloadVersion = 1;
    private static readonly JsonSerializerOptions JsonOptions = new(JsonSerializerDefaults.Web);
    private readonly string _connectionString;

    public HistoryStore(string databasePath)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(databasePath);
        var fullPath = Path.GetFullPath(databasePath);
        Directory.CreateDirectory(Path.GetDirectoryName(fullPath)!);
        _connectionString = new SqliteConnectionStringBuilder
        {
            DataSource = fullPath,
            Mode = SqliteOpenMode.ReadWriteCreate,
            Cache = SqliteCacheMode.Shared,
            ForeignKeys = true,
            Pooling = true
        }.ToString();
    }

    public async Task InitializeAsync(CancellationToken cancellationToken = default)
    {
        await using var connection = await OpenAsync(cancellationToken).ConfigureAwait(false);
        await ExecuteAsync(connection, "PRAGMA journal_mode=WAL;", cancellationToken).ConfigureAwait(false);
        await ExecuteAsync(connection, "PRAGMA synchronous=FULL;", cancellationToken).ConfigureAwait(false);
        await ExecuteAsync(connection, "PRAGMA foreign_keys=ON;", cancellationToken).ConfigureAwait(false);
        await ExecuteAsync(connection, "PRAGMA busy_timeout=5000;", cancellationToken).ConfigureAwait(false);

        var version = await GetSchemaVersionAsync(connection, cancellationToken).ConfigureAwait(false);
        if (version > CurrentSchemaVersion)
            throw new InvalidOperationException($"History database schema {version} is newer than this app supports ({CurrentSchemaVersion}).");
        if (version < 1)
            await ApplyMigration1Async(connection, cancellationToken).ConfigureAwait(false);
    }

    public async Task<TrackedWorkbook> AddBaselineAsync(
        string path,
        string reportDirectory,
        string sha256,
        DateTime fileLastWriteUtc,
        WorkbookSnapshot snapshot,
        CancellationToken cancellationToken = default)
    {
        ValidateSha256(sha256);
        var now = DateTime.UtcNow;
        var workbook = new TrackedWorkbook
        {
            Id = Guid.NewGuid(),
            Path = Path.GetFullPath(path),
            ReportDirectory = Path.GetFullPath(reportDirectory),
            IsEnabled = true,
            Status = TrackingStatus.Active,
            CreatedUtc = now,
            LastSuccessfulCaptureUtc = now,
            CurrentSequence = 0,
            CurrentHash = sha256,
            LastSummary = $"Baseline ready: {snapshot.Sheets.Count:N0} sheets, {snapshot.PopulatedCellCount:N0} populated cells"
        };

        await using var connection = await OpenAsync(cancellationToken).ConfigureAwait(false);
        await using var transaction = connection.BeginTransaction();
        try
        {
            await InsertWorkbookAsync(connection, transaction, workbook, cancellationToken).ConfigureAwait(false);
            await ReplaceSnapshotAsync(connection, transaction, workbook.Id, snapshot, cancellationToken).ConfigureAwait(false);
            await transaction.CommitAsync(cancellationToken).ConfigureAwait(false);
            return workbook;
        }
        catch
        {
            await transaction.RollbackAsync(CancellationToken.None).ConfigureAwait(false);
            throw;
        }
    }

    public async Task<VersionRecord> CommitVersionAsync(
        TrackedWorkbook expectedWorkbook,
        WorkbookSnapshot snapshot,
        WorkbookDiff diff,
        string sha256,
        DateTime fileLastWriteUtc,
        string reportPath,
        string summary,
        CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(expectedWorkbook);
        ArgumentNullException.ThrowIfNull(snapshot);
        ArgumentNullException.ThrowIfNull(diff);
        ValidateSha256(sha256);
        var capturedUtc = DateTime.UtcNow;

        await using var connection = await OpenAsync(cancellationToken).ConfigureAwait(false);
        await using var transaction = connection.BeginTransaction();
        try
        {
            var state = await ReadCommitStateAsync(connection, transaction, expectedWorkbook.Id, cancellationToken).ConfigureAwait(false);
            if (!state.Enabled)
                throw new InvalidOperationException("Tracking was paused before this capture could be committed.");
            if (!string.Equals(state.Hash, expectedWorkbook.CurrentHash, StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("The workbook baseline changed during capture; the save will be reconciled again.");
            if (string.Equals(state.Hash, sha256, StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("This workbook content is already the current version.");

            var sequence = state.Sequence + 1;
            var fullReportPath = Path.GetFullPath(reportPath);
            var versionId = await InsertVersionAsync(
                connection,
                transaction,
                expectedWorkbook.Id,
                sequence,
                capturedUtc,
                fileLastWriteUtc,
                state.Hash,
                sha256,
                diff,
                fullReportPath,
                summary,
                cancellationToken).ConfigureAwait(false);

            await InsertChangesAsync(connection, transaction, versionId, diff, cancellationToken).ConfigureAwait(false);
            await ReplaceSnapshotAsync(connection, transaction, expectedWorkbook.Id, snapshot, cancellationToken).ConfigureAwait(false);

            await using var update = connection.CreateCommand();
            update.Transaction = transaction;
            update.CommandText = """
                UPDATE tracked_workbooks
                SET current_sequence=$sequence, current_hash=$hash, last_success_utc=$last_success,
                    status=$status, last_summary=$summary, last_error=NULL
                WHERE id=$id AND is_enabled=1 AND current_sequence=$expected_sequence AND current_hash=$expected_hash;
                """;
            update.Parameters.AddWithValue("$sequence", sequence);
            update.Parameters.AddWithValue("$hash", sha256);
            update.Parameters.AddWithValue("$last_success", FormatDate(capturedUtc));
            update.Parameters.AddWithValue("$status", TrackingStatus.Active.ToString());
            update.Parameters.AddWithValue("$summary", summary);
            update.Parameters.AddWithValue("$id", expectedWorkbook.Id.ToString("D"));
            update.Parameters.AddWithValue("$expected_sequence", state.Sequence);
            update.Parameters.AddWithValue("$expected_hash", state.Hash);
            if (await update.ExecuteNonQueryAsync(cancellationToken).ConfigureAwait(false) != 1)
                throw new InvalidOperationException("The workbook state changed before the capture could be committed.");

            await transaction.CommitAsync(cancellationToken).ConfigureAwait(false);
            return new VersionRecord
            {
                Id = versionId,
                WorkbookId = expectedWorkbook.Id,
                WorkbookPath = expectedWorkbook.Path,
                Sequence = sequence,
                CapturedUtc = capturedUtc,
                FileLastWriteUtc = fileLastWriteUtc,
                PreviousSha256 = state.Hash,
                Sha256 = sha256,
                Status = VersionStatus.Captured,
                ReportStatus = ReportStatus.Pending,
                CellChangeCount = diff.CellChanges.Count,
                SheetChangeCount = diff.SheetChanges.Count,
                LiteralChangeCount = diff.LiteralChangeCount,
                FormulaChangeCount = diff.FormulaChangeCount,
                FormulaResultChangeCount = diff.FormulaResultChangeCount,
                CellTypeChangeCount = diff.CellTypeChangeCount,
                ReportPath = fullReportPath,
                Summary = summary
            };
        }
        catch
        {
            await transaction.RollbackAsync(CancellationToken.None).ConfigureAwait(false);
            throw;
        }
    }

    public async Task RecordErrorAsync(
        Guid workbookId,
        string fingerprint,
        string? sha256,
        string category,
        string message,
        TrackingStatus status,
        CancellationToken cancellationToken = default)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(fingerprint);
        ArgumentException.ThrowIfNullOrWhiteSpace(category);
        ArgumentException.ThrowIfNullOrWhiteSpace(message);
        var now = FormatDate(DateTime.UtcNow);
        await using var connection = await OpenAsync(cancellationToken).ConfigureAwait(false);
        await using var transaction = connection.BeginTransaction();

        await using (var command = connection.CreateCommand())
        {
            command.Transaction = transaction;
            command.CommandText = """
                INSERT INTO capture_errors(workbook_id,fingerprint,first_seen_utc,last_seen_utc,occurrence_count,sha256,category,message)
                VALUES($workbook_id,$fingerprint,$now,$now,1,$sha256,$category,$message)
                ON CONFLICT(workbook_id,fingerprint) DO UPDATE SET
                    last_seen_utc=excluded.last_seen_utc,
                    occurrence_count=capture_errors.occurrence_count+1,
                    sha256=excluded.sha256, category=excluded.category, message=excluded.message;
                """;
            command.Parameters.AddWithValue("$workbook_id", workbookId.ToString("D"));
            command.Parameters.AddWithValue("$fingerprint", fingerprint);
            command.Parameters.AddWithValue("$now", now);
            command.Parameters.AddWithValue("$sha256", (object?)sha256 ?? DBNull.Value);
            command.Parameters.AddWithValue("$category", category);
            command.Parameters.AddWithValue("$message", message);
            await command.ExecuteNonQueryAsync(cancellationToken).ConfigureAwait(false);
        }

        await using (var update = connection.CreateCommand())
        {
            update.Transaction = transaction;
            update.CommandText = "UPDATE tracked_workbooks SET status=$status,last_error=$error WHERE id=$id AND is_enabled=1;";
            update.Parameters.AddWithValue("$status", status.ToString());
            update.Parameters.AddWithValue("$error", message);
            update.Parameters.AddWithValue("$id", workbookId.ToString("D"));
            await update.ExecuteNonQueryAsync(cancellationToken).ConfigureAwait(false);
        }
        await transaction.CommitAsync(cancellationToken).ConfigureAwait(false);
    }

    public async Task<IReadOnlyList<TrackedWorkbook>> GetTrackedWorkbooksAsync(CancellationToken cancellationToken = default)
    {
        var result = new List<TrackedWorkbook>();
        await using var connection = await OpenAsync(cancellationToken).ConfigureAwait(false);
        await using var command = connection.CreateCommand();
        command.CommandText = "SELECT * FROM tracked_workbooks ORDER BY path COLLATE NOCASE;";
        await using var reader = await command.ExecuteReaderAsync(cancellationToken).ConfigureAwait(false);
        while (await reader.ReadAsync(cancellationToken).ConfigureAwait(false))
            result.Add(ReadTrackedWorkbook(reader));
        return result;
    }

    public async Task<TrackedWorkbook?> GetTrackedWorkbookAsync(Guid id, CancellationToken cancellationToken = default)
    {
        await using var connection = await OpenAsync(cancellationToken).ConfigureAwait(false);
        await using var command = connection.CreateCommand();
        command.CommandText = "SELECT * FROM tracked_workbooks WHERE id=$id;";
        command.Parameters.AddWithValue("$id", id.ToString("D"));
        await using var reader = await command.ExecuteReaderAsync(cancellationToken).ConfigureAwait(false);
        return await reader.ReadAsync(cancellationToken).ConfigureAwait(false) ? ReadTrackedWorkbook(reader) : null;
    }

    public async Task<WorkbookSnapshot> LoadCurrentSnapshotAsync(TrackedWorkbook workbook, CancellationToken cancellationToken = default)
    {
        var sheets = new Dictionary<uint, SheetState>();
        var cellsBySheet = new Dictionary<uint, Dictionary<string, CellState>>();
        await using var connection = await OpenAsync(cancellationToken).ConfigureAwait(false);
        await using var transaction = connection.BeginTransaction();

        await using (var command = connection.CreateCommand())
        {
            command.Transaction = transaction;
            command.CommandText = "SELECT sheet_id,state_json,payload_version FROM current_sheets WHERE workbook_id=$id;";
            command.Parameters.AddWithValue("$id", workbook.Id.ToString("D"));
            await using var reader = await command.ExecuteReaderAsync(cancellationToken).ConfigureAwait(false);
            while (await reader.ReadAsync(cancellationToken).ConfigureAwait(false))
            {
                EnsurePayloadVersion(reader.GetInt32(2));
                var sheet = Deserialize<SheetState>(reader.GetString(1));
                var sheetId = checked((uint)reader.GetInt64(0));
                sheets.Add(sheetId, sheet with { Cells = EmptyCells() });
                cellsBySheet.Add(sheetId, new Dictionary<string, CellState>(StringComparer.OrdinalIgnoreCase));
            }
        }

        await using (var command = connection.CreateCommand())
        {
            command.Transaction = transaction;
            command.CommandText = "SELECT sheet_id,address,state_json,payload_version FROM current_cells WHERE workbook_id=$id;";
            command.Parameters.AddWithValue("$id", workbook.Id.ToString("D"));
            await using var reader = await command.ExecuteReaderAsync(cancellationToken).ConfigureAwait(false);
            while (await reader.ReadAsync(cancellationToken).ConfigureAwait(false))
            {
                EnsurePayloadVersion(reader.GetInt32(3));
                var sheetId = checked((uint)reader.GetInt64(0));
                if (!cellsBySheet.TryGetValue(sheetId, out var cells))
                    throw new InvalidDataException("Stored cell references a missing sheet.");
                cells.Add(reader.GetString(1), Deserialize<CellState>(reader.GetString(2)));
            }
        }

        await transaction.CommitAsync(cancellationToken).ConfigureAwait(false);
        foreach (var sheetId in sheets.Keys.ToList())
            sheets[sheetId] = sheets[sheetId] with { Cells = new ReadOnlyDictionary<string, CellState>(cellsBySheet[sheetId]) };

        return new WorkbookSnapshot
        {
            SourcePath = workbook.Path,
            CapturedAtUtc = workbook.LastSuccessfulCaptureUtc ?? workbook.CreatedUtc,
            Sheets = new ReadOnlyDictionary<uint, SheetState>(sheets)
        };
    }

    public async Task<IReadOnlyList<VersionRecord>> GetVersionsAsync(Guid? workbookId = null, int limit = 500, CancellationToken cancellationToken = default)
    {
        var result = new List<VersionRecord>();
        await using var connection = await OpenAsync(cancellationToken).ConfigureAwait(false);
        await using var command = connection.CreateCommand();
        command.CommandText = workbookId is null
            ? "SELECT v.*,w.path FROM versions v JOIN tracked_workbooks w ON w.id=v.workbook_id ORDER BY v.captured_utc DESC LIMIT $limit;"
            : "SELECT v.*,w.path FROM versions v JOIN tracked_workbooks w ON w.id=v.workbook_id WHERE v.workbook_id=$workbook_id ORDER BY v.captured_utc DESC LIMIT $limit;";
        command.Parameters.AddWithValue("$limit", Math.Clamp(limit, 1, 10_000));
        if (workbookId is not null)
            command.Parameters.AddWithValue("$workbook_id", workbookId.Value.ToString("D"));
        await using var reader = await command.ExecuteReaderAsync(cancellationToken).ConfigureAwait(false);
        while (await reader.ReadAsync(cancellationToken).ConfigureAwait(false))
            result.Add(ReadVersion(reader));
        return result;
    }

    public async Task<IReadOnlyList<CaptureErrorRecord>> GetErrorsAsync(Guid? workbookId = null, int limit = 500, CancellationToken cancellationToken = default)
    {
        var result = new List<CaptureErrorRecord>();
        await using var connection = await OpenAsync(cancellationToken).ConfigureAwait(false);
        await using var command = connection.CreateCommand();
        command.CommandText = workbookId is null
            ? "SELECT e.*,w.path FROM capture_errors e JOIN tracked_workbooks w ON w.id=e.workbook_id ORDER BY e.last_seen_utc DESC LIMIT $limit;"
            : "SELECT e.*,w.path FROM capture_errors e JOIN tracked_workbooks w ON w.id=e.workbook_id WHERE e.workbook_id=$workbook_id ORDER BY e.last_seen_utc DESC LIMIT $limit;";
        command.Parameters.AddWithValue("$limit", Math.Clamp(limit, 1, 10_000));
        if (workbookId is not null)
            command.Parameters.AddWithValue("$workbook_id", workbookId.Value.ToString("D"));
        await using var reader = await command.ExecuteReaderAsync(cancellationToken).ConfigureAwait(false);
        while (await reader.ReadAsync(cancellationToken).ConfigureAwait(false))
        {
            result.Add(new CaptureErrorRecord
            {
                Id = reader.GetInt64(reader.GetOrdinal("id")),
                WorkbookId = Guid.Parse(reader.GetString(reader.GetOrdinal("workbook_id"))),
                WorkbookPath = reader.GetString(reader.GetOrdinal("path")),
                FirstSeenUtc = ParseDate(reader.GetString(reader.GetOrdinal("first_seen_utc"))),
                LastSeenUtc = ParseDate(reader.GetString(reader.GetOrdinal("last_seen_utc"))),
                OccurrenceCount = reader.GetInt32(reader.GetOrdinal("occurrence_count")),
                Sha256 = GetNullableString(reader, "sha256"),
                Category = reader.GetString(reader.GetOrdinal("category")),
                Message = reader.GetString(reader.GetOrdinal("message"))
            });
        }
        return result;
    }

    public async Task<IReadOnlyList<VersionRecord>> GetPendingReportsAsync(Guid? workbookId = null, CancellationToken cancellationToken = default)
    {
        var result = new List<VersionRecord>();
        await using var connection = await OpenAsync(cancellationToken).ConfigureAwait(false);
        await using var command = connection.CreateCommand();
        command.CommandText = workbookId is null
            ? "SELECT v.*,w.path FROM versions v JOIN tracked_workbooks w ON w.id=v.workbook_id WHERE v.report_status=$status ORDER BY v.captured_utc;"
            : "SELECT v.*,w.path FROM versions v JOIN tracked_workbooks w ON w.id=v.workbook_id WHERE v.report_status=$status AND v.workbook_id=$workbook_id ORDER BY v.captured_utc;";
        command.Parameters.AddWithValue("$status", ReportStatus.Pending.ToString());
        if (workbookId is not null)
            command.Parameters.AddWithValue("$workbook_id", workbookId.Value.ToString("D"));
        await using var reader = await command.ExecuteReaderAsync(cancellationToken).ConfigureAwait(false);
        while (await reader.ReadAsync(cancellationToken).ConfigureAwait(false))
            result.Add(ReadVersion(reader));
        return result;
    }

    public async Task MarkReportReadyAsync(long versionId, CancellationToken cancellationToken = default)
    {
        await using var connection = await OpenAsync(cancellationToken).ConfigureAwait(false);
        await using var command = connection.CreateCommand();
        command.CommandText = "UPDATE versions SET report_status=$status WHERE id=$id;";
        command.Parameters.AddWithValue("$status", ReportStatus.Ready.ToString());
        command.Parameters.AddWithValue("$id", versionId);
        if (await command.ExecuteNonQueryAsync(cancellationToken).ConfigureAwait(false) != 1)
            throw new KeyNotFoundException($"Version {versionId} does not exist.");
    }

    public async Task<IReadOnlyList<string>> GetReportPathsAsync(Guid workbookId, CancellationToken cancellationToken = default)
    {
        var result = new List<string>();
        await using var connection = await OpenAsync(cancellationToken).ConfigureAwait(false);
        await using var command = connection.CreateCommand();
        command.CommandText = "SELECT report_path FROM versions WHERE workbook_id=$id ORDER BY sequence;";
        command.Parameters.AddWithValue("$id", workbookId.ToString("D"));
        await using var reader = await command.ExecuteReaderAsync(cancellationToken).ConfigureAwait(false);
        while (await reader.ReadAsync(cancellationToken).ConfigureAwait(false))
            result.Add(reader.GetString(0));
        return result;
    }

    public async Task MarkReportsPendingAsync(Guid workbookId, CancellationToken cancellationToken = default)
    {
        await using var connection = await OpenAsync(cancellationToken).ConfigureAwait(false);
        await using var command = connection.CreateCommand();
        command.CommandText = "UPDATE versions SET report_status=$status WHERE workbook_id=$id;";
        command.Parameters.AddWithValue("$status", ReportStatus.Pending.ToString());
        command.Parameters.AddWithValue("$id", workbookId.ToString("D"));
        await command.ExecuteNonQueryAsync(cancellationToken).ConfigureAwait(false);
    }

    public async Task<WorkbookDiff> GetVersionDiffAsync(long versionId, CancellationToken cancellationToken = default)
    {
        var cells = new List<CellDelta>();
        var sheets = new List<SheetDelta>();
        await using var connection = await OpenAsync(cancellationToken).ConfigureAwait(false);
        await using var transaction = connection.BeginTransaction();

        await using (var exists = connection.CreateCommand())
        {
            exists.Transaction = transaction;
            exists.CommandText = "SELECT COUNT(*) FROM versions WHERE id=$id;";
            exists.Parameters.AddWithValue("$id", versionId);
            if (Convert.ToInt32(await exists.ExecuteScalarAsync(cancellationToken).ConfigureAwait(false), CultureInfo.InvariantCulture) != 1)
                throw new KeyNotFoundException($"Version {versionId} does not exist.");
        }

        await using (var command = connection.CreateCommand())
        {
            command.Transaction = transaction;
            command.CommandText = "SELECT sheet_id,sheet_name,address,kinds,before_json,after_json,payload_version FROM cell_changes WHERE version_id=$id ORDER BY id;";
            command.Parameters.AddWithValue("$id", versionId);
            await using var reader = await command.ExecuteReaderAsync(cancellationToken).ConfigureAwait(false);
            while (await reader.ReadAsync(cancellationToken).ConfigureAwait(false))
            {
                EnsurePayloadVersion(reader.GetInt32(6));
                cells.Add(new CellDelta
                {
                    SheetId = checked((uint)reader.GetInt64(0)),
                    SheetName = reader.GetString(1),
                    Address = reader.GetString(2),
                    Kinds = reader.GetString(3).Split(',', StringSplitOptions.RemoveEmptyEntries).Select(value => Enum.Parse<CellChangeKind>(value)).ToArray(),
                    Before = DeserializeNullable<CellState>(reader, 4),
                    After = DeserializeNullable<CellState>(reader, 5)
                });
            }
        }

        await using (var command = connection.CreateCommand())
        {
            command.Transaction = transaction;
            command.CommandText = "SELECT sheet_id,kind,before_json,after_json,payload_version FROM sheet_changes WHERE version_id=$id ORDER BY id;";
            command.Parameters.AddWithValue("$id", versionId);
            await using var reader = await command.ExecuteReaderAsync(cancellationToken).ConfigureAwait(false);
            while (await reader.ReadAsync(cancellationToken).ConfigureAwait(false))
            {
                EnsurePayloadVersion(reader.GetInt32(4));
                sheets.Add(new SheetDelta
                {
                    SheetId = checked((uint)reader.GetInt64(0)),
                    Kind = Enum.Parse<SheetChangeKind>(reader.GetString(1)),
                    Before = DeserializeNullable<SheetState>(reader, 2),
                    After = DeserializeNullable<SheetState>(reader, 3)
                });
            }
        }

        await transaction.CommitAsync(cancellationToken).ConfigureAwait(false);
        return new WorkbookDiff { CellChanges = cells, SheetChanges = sheets };
    }

    public async Task SetEnabledAsync(Guid workbookId, bool enabled, CancellationToken cancellationToken = default)
    {
        await using var connection = await OpenAsync(cancellationToken).ConfigureAwait(false);
        await using var command = connection.CreateCommand();
        command.CommandText = "UPDATE tracked_workbooks SET is_enabled=$enabled,status=$status WHERE id=$id;";
        command.Parameters.AddWithValue("$enabled", enabled ? 1 : 0);
        command.Parameters.AddWithValue("$status", enabled ? TrackingStatus.Active.ToString() : TrackingStatus.Paused.ToString());
        command.Parameters.AddWithValue("$id", workbookId.ToString("D"));
        await command.ExecuteNonQueryAsync(cancellationToken).ConfigureAwait(false);
    }

    public async Task UpdateReportDirectoryAsync(Guid workbookId, string reportDirectory, CancellationToken cancellationToken = default)
    {
        await using var connection = await OpenAsync(cancellationToken).ConfigureAwait(false);
        await using var command = connection.CreateCommand();
        command.CommandText = "UPDATE tracked_workbooks SET report_directory=$directory WHERE id=$id;";
        command.Parameters.AddWithValue("$directory", Path.GetFullPath(reportDirectory));
        command.Parameters.AddWithValue("$id", workbookId.ToString("D"));
        await command.ExecuteNonQueryAsync(cancellationToken).ConfigureAwait(false);
    }

    public Task StopTrackingAsync(Guid workbookId, CancellationToken cancellationToken = default) => SetEnabledAsync(workbookId, false, cancellationToken);

    public async Task<IReadOnlyList<string>> PurgeAsync(Guid workbookId, bool confirmed, CancellationToken cancellationToken = default)
    {
        if (!confirmed)
            throw new InvalidOperationException("Permanent deletion requires explicit confirmation.");
        var reports = new List<string>();
        await using var connection = await OpenAsync(cancellationToken).ConfigureAwait(false);
        await using var transaction = connection.BeginTransaction();
        await using (var paths = connection.CreateCommand())
        {
            paths.Transaction = transaction;
            paths.CommandText = "SELECT report_path FROM versions WHERE workbook_id=$id;";
            paths.Parameters.AddWithValue("$id", workbookId.ToString("D"));
            await using var reader = await paths.ExecuteReaderAsync(cancellationToken).ConfigureAwait(false);
            while (await reader.ReadAsync(cancellationToken).ConfigureAwait(false))
                reports.Add(reader.GetString(0));
        }
        await using (var delete = connection.CreateCommand())
        {
            delete.Transaction = transaction;
            delete.CommandText = "DELETE FROM tracked_workbooks WHERE id=$id;";
            delete.Parameters.AddWithValue("$id", workbookId.ToString("D"));
            await delete.ExecuteNonQueryAsync(cancellationToken).ConfigureAwait(false);
        }
        await transaction.CommitAsync(cancellationToken).ConfigureAwait(false);
        return reports;
    }

    public async Task SetProcessingStatusAsync(Guid workbookId, TrackingStatus status, string? error = null, CancellationToken cancellationToken = default)
    {
        await using var connection = await OpenAsync(cancellationToken).ConfigureAwait(false);
        await using var command = connection.CreateCommand();
        command.CommandText = "UPDATE tracked_workbooks SET status=$status,last_error=$error WHERE id=$id AND is_enabled=1;";
        command.Parameters.AddWithValue("$status", status.ToString());
        command.Parameters.AddWithValue("$error", (object?)error ?? DBNull.Value);
        command.Parameters.AddWithValue("$id", workbookId.ToString("D"));
        await command.ExecuteNonQueryAsync(cancellationToken).ConfigureAwait(false);
    }

    public async Task<string?> GetSettingAsync(string key, CancellationToken cancellationToken = default)
    {
        await using var connection = await OpenAsync(cancellationToken).ConfigureAwait(false);
        await using var command = connection.CreateCommand();
        command.CommandText = "SELECT value FROM app_settings WHERE key=$key;";
        command.Parameters.AddWithValue("$key", key);
        return await command.ExecuteScalarAsync(cancellationToken).ConfigureAwait(false) as string;
    }

    public async Task SetSettingAsync(string key, string value, CancellationToken cancellationToken = default)
    {
        await using var connection = await OpenAsync(cancellationToken).ConfigureAwait(false);
        await using var command = connection.CreateCommand();
        command.CommandText = "INSERT INTO app_settings(key,value) VALUES($key,$value) ON CONFLICT(key) DO UPDATE SET value=excluded.value;";
        command.Parameters.AddWithValue("$key", key);
        command.Parameters.AddWithValue("$value", value);
        await command.ExecuteNonQueryAsync(cancellationToken).ConfigureAwait(false);
    }

    private static async Task ApplyMigration1Async(SqliteConnection connection, CancellationToken cancellationToken)
    {
        await using var transaction = connection.BeginTransaction();
        await using var command = connection.CreateCommand();
        command.Transaction = transaction;
        command.CommandText = """
            CREATE TABLE schema_migrations(version INTEGER PRIMARY KEY CHECK(version>0),applied_utc TEXT NOT NULL);
            CREATE TABLE app_settings(key TEXT PRIMARY KEY,value TEXT NOT NULL);
            CREATE TABLE tracked_workbooks(
                id TEXT PRIMARY KEY,path TEXT NOT NULL UNIQUE COLLATE NOCASE,report_directory TEXT NOT NULL,
                is_enabled INTEGER NOT NULL CHECK(is_enabled IN(0,1)),status TEXT NOT NULL,created_utc TEXT NOT NULL,
                last_success_utc TEXT,current_sequence INTEGER NOT NULL CHECK(current_sequence>=0),
                current_hash TEXT NOT NULL CHECK(length(current_hash)=64),last_summary TEXT,last_error TEXT);
            CREATE TABLE current_sheets(
                workbook_id TEXT NOT NULL,sheet_id INTEGER NOT NULL,state_json TEXT NOT NULL,payload_version INTEGER NOT NULL,
                PRIMARY KEY(workbook_id,sheet_id),FOREIGN KEY(workbook_id) REFERENCES tracked_workbooks(id) ON DELETE CASCADE);
            CREATE TABLE current_cells(
                workbook_id TEXT NOT NULL,sheet_id INTEGER NOT NULL,address TEXT NOT NULL COLLATE NOCASE,state_json TEXT NOT NULL,payload_version INTEGER NOT NULL,
                PRIMARY KEY(workbook_id,sheet_id,address),FOREIGN KEY(workbook_id,sheet_id) REFERENCES current_sheets(workbook_id,sheet_id) ON DELETE CASCADE);
            CREATE TABLE versions(
                id INTEGER PRIMARY KEY AUTOINCREMENT,workbook_id TEXT NOT NULL,sequence INTEGER NOT NULL CHECK(sequence>0),
                captured_utc TEXT NOT NULL,file_last_write_utc TEXT NOT NULL,previous_sha256 TEXT NOT NULL CHECK(length(previous_sha256)=64),
                sha256 TEXT NOT NULL CHECK(length(sha256)=64),status TEXT NOT NULL,report_status TEXT NOT NULL,cell_change_count INTEGER NOT NULL CHECK(cell_change_count>=0),
                sheet_change_count INTEGER NOT NULL CHECK(sheet_change_count>=0),literal_change_count INTEGER NOT NULL CHECK(literal_change_count>=0),
                formula_change_count INTEGER NOT NULL CHECK(formula_change_count>=0),formula_result_change_count INTEGER NOT NULL CHECK(formula_result_change_count>=0),
                cell_type_change_count INTEGER NOT NULL CHECK(cell_type_change_count>=0),report_path TEXT NOT NULL,summary TEXT,
                UNIQUE(workbook_id,sequence),FOREIGN KEY(workbook_id) REFERENCES tracked_workbooks(id) ON DELETE CASCADE);
            CREATE TABLE cell_changes(
                id INTEGER PRIMARY KEY AUTOINCREMENT,version_id INTEGER NOT NULL,sheet_id INTEGER NOT NULL,sheet_name TEXT NOT NULL,
                address TEXT NOT NULL,kinds TEXT NOT NULL,before_json TEXT,after_json TEXT,payload_version INTEGER NOT NULL,
                FOREIGN KEY(version_id) REFERENCES versions(id) ON DELETE CASCADE);
            CREATE TABLE sheet_changes(
                id INTEGER PRIMARY KEY AUTOINCREMENT,version_id INTEGER NOT NULL,sheet_id INTEGER NOT NULL,kind TEXT NOT NULL,
                before_json TEXT,after_json TEXT,payload_version INTEGER NOT NULL,
                FOREIGN KEY(version_id) REFERENCES versions(id) ON DELETE CASCADE);
            CREATE TABLE capture_errors(
                id INTEGER PRIMARY KEY AUTOINCREMENT,workbook_id TEXT NOT NULL,fingerprint TEXT NOT NULL,first_seen_utc TEXT NOT NULL,
                last_seen_utc TEXT NOT NULL,occurrence_count INTEGER NOT NULL CHECK(occurrence_count>0),sha256 TEXT,category TEXT NOT NULL,message TEXT NOT NULL,
                UNIQUE(workbook_id,fingerprint),FOREIGN KEY(workbook_id) REFERENCES tracked_workbooks(id) ON DELETE CASCADE);
            CREATE INDEX ix_versions_workbook_captured ON versions(workbook_id,captured_utc DESC);
            CREATE INDEX ix_cell_changes_version ON cell_changes(version_id);
            CREATE INDEX ix_sheet_changes_version ON sheet_changes(version_id);
            CREATE INDEX ix_capture_errors_workbook_seen ON capture_errors(workbook_id,last_seen_utc DESC);
            INSERT INTO schema_migrations(version,applied_utc) VALUES(1,$applied_utc);
            PRAGMA user_version=1;
            """;
        command.Parameters.AddWithValue("$applied_utc", FormatDate(DateTime.UtcNow));
        await command.ExecuteNonQueryAsync(cancellationToken).ConfigureAwait(false);
        await transaction.CommitAsync(cancellationToken).ConfigureAwait(false);
    }

    private static async Task<int> GetSchemaVersionAsync(SqliteConnection connection, CancellationToken cancellationToken)
    {
        await using var command = connection.CreateCommand();
        command.CommandText = "PRAGMA user_version;";
        return Convert.ToInt32(await command.ExecuteScalarAsync(cancellationToken).ConfigureAwait(false), CultureInfo.InvariantCulture);
    }

    private async Task<SqliteConnection> OpenAsync(CancellationToken cancellationToken)
    {
        var connection = new SqliteConnection(_connectionString);
        await connection.OpenAsync(cancellationToken).ConfigureAwait(false);
        return connection;
    }

    private static async Task ExecuteAsync(SqliteConnection connection, string sql, CancellationToken cancellationToken)
    {
        await using var command = connection.CreateCommand();
        command.CommandText = sql;
        await command.ExecuteNonQueryAsync(cancellationToken).ConfigureAwait(false);
    }

    private static async Task InsertWorkbookAsync(SqliteConnection connection, SqliteTransaction transaction, TrackedWorkbook workbook, CancellationToken cancellationToken)
    {
        await using var command = connection.CreateCommand();
        command.Transaction = transaction;
        command.CommandText = """
            INSERT INTO tracked_workbooks(id,path,report_directory,is_enabled,status,created_utc,last_success_utc,current_sequence,current_hash,last_summary,last_error)
            VALUES($id,$path,$report_directory,$enabled,$status,$created,$success,$sequence,$hash,$summary,$error);
            """;
        command.Parameters.AddWithValue("$id", workbook.Id.ToString("D"));
        command.Parameters.AddWithValue("$path", workbook.Path);
        command.Parameters.AddWithValue("$report_directory", workbook.ReportDirectory);
        command.Parameters.AddWithValue("$enabled", workbook.IsEnabled ? 1 : 0);
        command.Parameters.AddWithValue("$status", workbook.Status.ToString());
        command.Parameters.AddWithValue("$created", FormatDate(workbook.CreatedUtc));
        command.Parameters.AddWithValue("$success", workbook.LastSuccessfulCaptureUtc is null ? DBNull.Value : FormatDate(workbook.LastSuccessfulCaptureUtc.Value));
        command.Parameters.AddWithValue("$sequence", workbook.CurrentSequence);
        command.Parameters.AddWithValue("$hash", workbook.CurrentHash);
        command.Parameters.AddWithValue("$summary", (object?)workbook.LastSummary ?? DBNull.Value);
        command.Parameters.AddWithValue("$error", (object?)workbook.LastError ?? DBNull.Value);
        await command.ExecuteNonQueryAsync(cancellationToken).ConfigureAwait(false);
    }

    private static async Task ReplaceSnapshotAsync(SqliteConnection connection, SqliteTransaction transaction, Guid workbookId, WorkbookSnapshot snapshot, CancellationToken cancellationToken)
    {
        await using (var clear = connection.CreateCommand())
        {
            clear.Transaction = transaction;
            clear.CommandText = "DELETE FROM current_sheets WHERE workbook_id=$id;";
            clear.Parameters.AddWithValue("$id", workbookId.ToString("D"));
            await clear.ExecuteNonQueryAsync(cancellationToken).ConfigureAwait(false);
        }

        await using var sheetCommand = connection.CreateCommand();
        sheetCommand.Transaction = transaction;
        sheetCommand.CommandText = "INSERT INTO current_sheets(workbook_id,sheet_id,state_json,payload_version) VALUES($workbook_id,$sheet_id,$json,$payload);";
        var sheetWorkbook = sheetCommand.Parameters.Add("$workbook_id", SqliteType.Text);
        var sheetId = sheetCommand.Parameters.Add("$sheet_id", SqliteType.Integer);
        var sheetJson = sheetCommand.Parameters.Add("$json", SqliteType.Text);
        sheetCommand.Parameters.AddWithValue("$payload", PayloadVersion);

        await using var cellCommand = connection.CreateCommand();
        cellCommand.Transaction = transaction;
        cellCommand.CommandText = "INSERT INTO current_cells(workbook_id,sheet_id,address,state_json,payload_version) VALUES($workbook_id,$sheet_id,$address,$json,$payload);";
        var cellWorkbook = cellCommand.Parameters.Add("$workbook_id", SqliteType.Text);
        var cellSheet = cellCommand.Parameters.Add("$sheet_id", SqliteType.Integer);
        var cellAddress = cellCommand.Parameters.Add("$address", SqliteType.Text);
        var cellJson = cellCommand.Parameters.Add("$json", SqliteType.Text);
        cellCommand.Parameters.AddWithValue("$payload", PayloadVersion);

        var workbookValue = workbookId.ToString("D");
        foreach (var sheet in snapshot.Sheets.Values)
        {
            sheetWorkbook.Value = workbookValue;
            sheetId.Value = sheet.SheetId;
            sheetJson.Value = JsonSerializer.Serialize(sheet with { Cells = EmptyCells() }, JsonOptions);
            await sheetCommand.ExecuteNonQueryAsync(cancellationToken).ConfigureAwait(false);
            foreach (var cell in sheet.Cells.Values)
            {
                cellWorkbook.Value = workbookValue;
                cellSheet.Value = sheet.SheetId;
                cellAddress.Value = cell.Address;
                cellJson.Value = JsonSerializer.Serialize(cell, JsonOptions);
                await cellCommand.ExecuteNonQueryAsync(cancellationToken).ConfigureAwait(false);
            }
        }
    }

    private static async Task<long> InsertVersionAsync(
        SqliteConnection connection, SqliteTransaction transaction, Guid workbookId, long sequence,
        DateTime capturedUtc, DateTime fileLastWriteUtc, string previousSha256, string sha256,
        WorkbookDiff diff, string reportPath, string summary, CancellationToken cancellationToken)
    {
        await using var command = connection.CreateCommand();
        command.Transaction = transaction;
        command.CommandText = """
            INSERT INTO versions(workbook_id,sequence,captured_utc,file_last_write_utc,previous_sha256,sha256,status,report_status,
                cell_change_count,sheet_change_count,literal_change_count,formula_change_count,formula_result_change_count,cell_type_change_count,report_path,summary)
            VALUES($workbook_id,$sequence,$captured,$last_write,$previous_hash,$hash,$status,$report_status,
                $cell_count,$sheet_count,$literal_count,$formula_count,$formula_result_count,$type_count,$report_path,$summary);
            SELECT last_insert_rowid();
            """;
        command.Parameters.AddWithValue("$workbook_id", workbookId.ToString("D"));
        command.Parameters.AddWithValue("$sequence", sequence);
        command.Parameters.AddWithValue("$captured", FormatDate(capturedUtc));
        command.Parameters.AddWithValue("$last_write", FormatDate(fileLastWriteUtc));
        command.Parameters.AddWithValue("$previous_hash", previousSha256);
        command.Parameters.AddWithValue("$hash", sha256);
        command.Parameters.AddWithValue("$status", VersionStatus.Captured.ToString());
        command.Parameters.AddWithValue("$report_status", ReportStatus.Pending.ToString());
        command.Parameters.AddWithValue("$cell_count", diff.CellChanges.Count);
        command.Parameters.AddWithValue("$sheet_count", diff.SheetChanges.Count);
        command.Parameters.AddWithValue("$literal_count", diff.LiteralChangeCount);
        command.Parameters.AddWithValue("$formula_count", diff.FormulaChangeCount);
        command.Parameters.AddWithValue("$formula_result_count", diff.FormulaResultChangeCount);
        command.Parameters.AddWithValue("$type_count", diff.CellTypeChangeCount);
        command.Parameters.AddWithValue("$report_path", reportPath);
        command.Parameters.AddWithValue("$summary", summary);
        return Convert.ToInt64(await command.ExecuteScalarAsync(cancellationToken).ConfigureAwait(false), CultureInfo.InvariantCulture);
    }

    private static async Task InsertChangesAsync(SqliteConnection connection, SqliteTransaction transaction, long versionId, WorkbookDiff diff, CancellationToken cancellationToken)
    {
        await using var cellCommand = connection.CreateCommand();
        cellCommand.Transaction = transaction;
        cellCommand.CommandText = "INSERT INTO cell_changes(version_id,sheet_id,sheet_name,address,kinds,before_json,after_json,payload_version) VALUES($version,$sheet,$name,$address,$kinds,$before,$after,$payload);";
        var cellVersion = cellCommand.Parameters.Add("$version", SqliteType.Integer);
        var cellSheet = cellCommand.Parameters.Add("$sheet", SqliteType.Integer);
        var cellName = cellCommand.Parameters.Add("$name", SqliteType.Text);
        var cellAddress = cellCommand.Parameters.Add("$address", SqliteType.Text);
        var cellKinds = cellCommand.Parameters.Add("$kinds", SqliteType.Text);
        var cellBefore = cellCommand.Parameters.Add("$before", SqliteType.Text);
        var cellAfter = cellCommand.Parameters.Add("$after", SqliteType.Text);
        cellCommand.Parameters.AddWithValue("$payload", PayloadVersion);
        foreach (var change in diff.CellChanges)
        {
            cellVersion.Value = versionId;
            cellSheet.Value = change.SheetId;
            cellName.Value = change.SheetName;
            cellAddress.Value = change.Address;
            cellKinds.Value = string.Join(',', change.Kinds);
            cellBefore.Value = change.Before is null ? DBNull.Value : JsonSerializer.Serialize(change.Before, JsonOptions);
            cellAfter.Value = change.After is null ? DBNull.Value : JsonSerializer.Serialize(change.After, JsonOptions);
            await cellCommand.ExecuteNonQueryAsync(cancellationToken).ConfigureAwait(false);
        }

        await using var sheetCommand = connection.CreateCommand();
        sheetCommand.Transaction = transaction;
        sheetCommand.CommandText = "INSERT INTO sheet_changes(version_id,sheet_id,kind,before_json,after_json,payload_version) VALUES($version,$sheet,$kind,$before,$after,$payload);";
        var sheetVersion = sheetCommand.Parameters.Add("$version", SqliteType.Integer);
        var sheetId = sheetCommand.Parameters.Add("$sheet", SqliteType.Integer);
        var sheetKind = sheetCommand.Parameters.Add("$kind", SqliteType.Text);
        var sheetBefore = sheetCommand.Parameters.Add("$before", SqliteType.Text);
        var sheetAfter = sheetCommand.Parameters.Add("$after", SqliteType.Text);
        sheetCommand.Parameters.AddWithValue("$payload", PayloadVersion);
        foreach (var change in diff.SheetChanges)
        {
            sheetVersion.Value = versionId;
            sheetId.Value = change.SheetId;
            sheetKind.Value = change.Kind.ToString();
            sheetBefore.Value = change.Before is null ? DBNull.Value : JsonSerializer.Serialize(change.Before with { Cells = EmptyCells() }, JsonOptions);
            sheetAfter.Value = change.After is null ? DBNull.Value : JsonSerializer.Serialize(change.After with { Cells = EmptyCells() }, JsonOptions);
            await sheetCommand.ExecuteNonQueryAsync(cancellationToken).ConfigureAwait(false);
        }
    }

    private static async Task<(long Sequence, string Hash, bool Enabled)> ReadCommitStateAsync(SqliteConnection connection, SqliteTransaction transaction, Guid workbookId, CancellationToken cancellationToken)
    {
        await using var command = connection.CreateCommand();
        command.Transaction = transaction;
        command.CommandText = "SELECT current_sequence,current_hash,is_enabled FROM tracked_workbooks WHERE id=$id;";
        command.Parameters.AddWithValue("$id", workbookId.ToString("D"));
        await using var reader = await command.ExecuteReaderAsync(cancellationToken).ConfigureAwait(false);
        if (!await reader.ReadAsync(cancellationToken).ConfigureAwait(false))
            throw new KeyNotFoundException("Tracked workbook no longer exists.");
        return (reader.GetInt64(0), reader.GetString(1), reader.GetInt32(2) != 0);
    }

    private static VersionRecord ReadVersion(SqliteDataReader reader) => new()
    {
        Id = reader.GetInt64(reader.GetOrdinal("id")),
        WorkbookId = Guid.Parse(reader.GetString(reader.GetOrdinal("workbook_id"))),
        WorkbookPath = reader.GetString(reader.GetOrdinal("path")),
        Sequence = reader.GetInt64(reader.GetOrdinal("sequence")),
        CapturedUtc = ParseDate(reader.GetString(reader.GetOrdinal("captured_utc"))),
        FileLastWriteUtc = ParseDate(reader.GetString(reader.GetOrdinal("file_last_write_utc"))),
        PreviousSha256 = reader.GetString(reader.GetOrdinal("previous_sha256")),
        Sha256 = reader.GetString(reader.GetOrdinal("sha256")),
        Status = Enum.Parse<VersionStatus>(reader.GetString(reader.GetOrdinal("status"))),
        ReportStatus = Enum.Parse<ReportStatus>(reader.GetString(reader.GetOrdinal("report_status"))),
        CellChangeCount = reader.GetInt32(reader.GetOrdinal("cell_change_count")),
        SheetChangeCount = reader.GetInt32(reader.GetOrdinal("sheet_change_count")),
        LiteralChangeCount = reader.GetInt32(reader.GetOrdinal("literal_change_count")),
        FormulaChangeCount = reader.GetInt32(reader.GetOrdinal("formula_change_count")),
        FormulaResultChangeCount = reader.GetInt32(reader.GetOrdinal("formula_result_change_count")),
        CellTypeChangeCount = reader.GetInt32(reader.GetOrdinal("cell_type_change_count")),
        ReportPath = GetNullableString(reader, "report_path"),
        Summary = GetNullableString(reader, "summary")
    };

    private static TrackedWorkbook ReadTrackedWorkbook(SqliteDataReader reader) => new()
    {
        Id = Guid.Parse(reader.GetString(reader.GetOrdinal("id"))),
        Path = reader.GetString(reader.GetOrdinal("path")),
        ReportDirectory = reader.GetString(reader.GetOrdinal("report_directory")),
        IsEnabled = reader.GetInt32(reader.GetOrdinal("is_enabled")) != 0,
        Status = Enum.Parse<TrackingStatus>(reader.GetString(reader.GetOrdinal("status"))),
        CreatedUtc = ParseDate(reader.GetString(reader.GetOrdinal("created_utc"))),
        LastSuccessfulCaptureUtc = GetNullableDate(reader, "last_success_utc"),
        CurrentSequence = reader.GetInt64(reader.GetOrdinal("current_sequence")),
        CurrentHash = reader.GetString(reader.GetOrdinal("current_hash")),
        LastSummary = GetNullableString(reader, "last_summary"),
        LastError = GetNullableString(reader, "last_error")
    };

    private static T Deserialize<T>(string json) => JsonSerializer.Deserialize<T>(json, JsonOptions)
        ?? throw new InvalidDataException($"Stored {typeof(T).Name} payload is invalid.");
    private static T? DeserializeNullable<T>(SqliteDataReader reader, int ordinal) where T : class => reader.IsDBNull(ordinal) ? null : Deserialize<T>(reader.GetString(ordinal));
    private static ReadOnlyDictionary<string, CellState> EmptyCells() => new(new Dictionary<string, CellState>(StringComparer.OrdinalIgnoreCase));
    private static void EnsurePayloadVersion(int version)
    {
        if (version != PayloadVersion)
            throw new InvalidDataException($"Unsupported stored payload version {version}.");
    }
    private static void ValidateSha256(string value)
    {
        if (value.Length != 64 || value.Any(character => !char.IsAsciiHexDigit(character)))
            throw new ArgumentException("SHA-256 must be 64 hexadecimal characters.", nameof(value));
    }
    private static string? GetNullableString(SqliteDataReader reader, string name)
    {
        var ordinal = reader.GetOrdinal(name);
        return reader.IsDBNull(ordinal) ? null : reader.GetString(ordinal);
    }
    private static DateTime? GetNullableDate(SqliteDataReader reader, string name) => GetNullableString(reader, name) is { } value ? ParseDate(value) : null;
    private static string FormatDate(DateTime value) => value.ToUniversalTime().ToString("O", CultureInfo.InvariantCulture);
    private static DateTime ParseDate(string value) => DateTime.Parse(value, CultureInfo.InvariantCulture, DateTimeStyles.RoundtripKind).ToUniversalTime();
}
