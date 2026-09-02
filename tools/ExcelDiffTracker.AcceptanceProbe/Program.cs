using System.Text.Json;
using Microsoft.Data.Sqlite;

namespace ExcelDiffTracker.AcceptanceProbe;

internal static class Program
{
    private static readonly JsonSerializerOptions JsonOptions = new(JsonSerializerDefaults.Web) { WriteIndented = true };

    public static async Task<int> Main(string[] args)
    {
        try
        {
            var options = Options.Parse(args);
            var result = await InspectAsync(options).ConfigureAwait(false);
            Console.WriteLine(JsonSerializer.Serialize(result, JsonOptions));
            return result.Passed ? 0 : 1;
        }
        catch (Exception exception)
        {
            Console.Error.WriteLine(JsonSerializer.Serialize(new
            {
                passed = false,
                error = exception.Message,
                type = exception.GetType().FullName
            }, JsonOptions));
            return 2;
        }
    }

    private static async Task<ProbeResult> InspectAsync(Options options)
    {
        var failures = new List<string>();
        var connectionString = new SqliteConnectionStringBuilder
        {
            DataSource = Path.GetFullPath(options.DatabasePath),
            Mode = SqliteOpenMode.ReadOnly,
            Cache = SqliteCacheMode.Shared,
            Pooling = false
        }.ToString();
        await using var connection = new SqliteConnection(connectionString);
        await connection.OpenAsync().ConfigureAwait(false);

        await using var workbookCommand = connection.CreateCommand();
        workbookCommand.CommandText = """
            SELECT id,status,current_sequence,current_hash,last_error,report_directory
            FROM tracked_workbooks WHERE path=$path COLLATE NOCASE;
            """;
        workbookCommand.Parameters.AddWithValue("$path", Path.GetFullPath(options.WorkbookPath));
        await using var workbookReader = await workbookCommand.ExecuteReaderAsync().ConfigureAwait(false);
        if (!await workbookReader.ReadAsync().ConfigureAwait(false))
        {
            return new ProbeResult(false, ["Tracked workbook was not found."], null, null, null, null, null, null, null, null, null, null, null);
        }

        var workbookId = workbookReader.GetString(0);
        var status = workbookReader.GetString(1);
        var currentSequence = workbookReader.GetInt64(2);
        var currentHash = workbookReader.GetString(3);
        var lastError = workbookReader.IsDBNull(4) ? null : workbookReader.GetString(4);
        var reportDirectory = workbookReader.GetString(5);
        await workbookReader.DisposeAsync().ConfigureAwait(false);

        if (options.ExpectedSequence is { } expectedSequence && currentSequence != expectedSequence)
            failures.Add($"Expected sequence {expectedSequence}, found {currentSequence}.");
        if (options.RequireActive && !string.Equals(status, "Active", StringComparison.Ordinal))
            failures.Add($"Expected workbook status Active, found {status}.");
        if (options.ExpectedStatus is not null && !string.Equals(status, options.ExpectedStatus, StringComparison.Ordinal))
            failures.Add($"Expected workbook status {options.ExpectedStatus}, found {status}.");
        if (options.RequireNoLastError && !string.IsNullOrWhiteSpace(lastError))
            failures.Add($"Expected no current workbook error, found: {lastError}");

        await using var errorCommand = connection.CreateCommand();
        errorCommand.CommandText = "SELECT COUNT(*) FROM capture_errors WHERE workbook_id=$id;";
        errorCommand.Parameters.AddWithValue("$id", workbookId);
        var errorCount = Convert.ToInt64(await errorCommand.ExecuteScalarAsync().ConfigureAwait(false));
        if (options.RequireNoErrors && errorCount != 0)
            failures.Add($"Expected zero capture errors, found {errorCount}.");
        if (options.MinimumErrors is { } minimumErrors && errorCount < minimumErrors)
            failures.Add($"Expected at least {minimumErrors} capture errors, found {errorCount}.");
        if (options.ExpectedErrorCount is { } expectedErrorCount && errorCount != expectedErrorCount)
            failures.Add($"Expected exactly {expectedErrorCount} capture errors, found {errorCount}.");

        await using var versionCountCommand = connection.CreateCommand();
        versionCountCommand.CommandText = "SELECT COUNT(*), COUNT(DISTINCT sha256) FROM versions WHERE workbook_id=$id;";
        versionCountCommand.Parameters.AddWithValue("$id", workbookId);
        await using var versionCountReader = await versionCountCommand.ExecuteReaderAsync().ConfigureAwait(false);
        await versionCountReader.ReadAsync().ConfigureAwait(false);
        var versionCount = versionCountReader.GetInt64(0);
        var distinctVersionHashCount = versionCountReader.GetInt64(1);
        await versionCountReader.DisposeAsync().ConfigureAwait(false);
        if (options.ExpectedVersionCount is { } expectedVersionCount && versionCount != expectedVersionCount)
            failures.Add($"Expected exactly {expectedVersionCount} versions, found {versionCount}.");
        if (options.RequireUniqueVersionHashes && distinctVersionHashCount != versionCount)
            failures.Add($"Expected every captured version hash to be unique; found {distinctVersionHashCount} unique hashes across {versionCount} versions.");
        if (options.RequireSourceHashMatch)
        {
            using var stream = new FileStream(
                Path.GetFullPath(options.WorkbookPath), FileMode.Open, FileAccess.Read,
                FileShare.ReadWrite | FileShare.Delete, 1024 * 1024, FileOptions.SequentialScan);
            var sourceHash = Convert.ToHexString(System.Security.Cryptography.SHA256.HashData(stream));
            if (!string.Equals(sourceHash, currentHash, StringComparison.OrdinalIgnoreCase))
                failures.Add($"Tracked hash {currentHash} does not match current workbook bytes {sourceHash}.");
        }

        await using var versionCommand = connection.CreateCommand();
        versionCommand.CommandText = """
            SELECT id,sequence,sha256,report_status,report_path,cell_change_count,sheet_change_count,summary
            FROM versions WHERE workbook_id=$id ORDER BY sequence DESC LIMIT 1;
            """;
        versionCommand.Parameters.AddWithValue("$id", workbookId);
        await using var versionReader = await versionCommand.ExecuteReaderAsync().ConfigureAwait(false);
        long? versionId = null;
        long? latestSequence = null;
        string? versionHash = null;
        string? reportStatus = null;
        string? reportPath = null;
        long? cellChangeCount = null;
        long? sheetChangeCount = null;
        string? summary = null;
        if (await versionReader.ReadAsync().ConfigureAwait(false))
        {
            versionId = versionReader.GetInt64(0);
            latestSequence = versionReader.GetInt64(1);
            versionHash = versionReader.GetString(2);
            reportStatus = versionReader.GetString(3);
            reportPath = versionReader.GetString(4);
            cellChangeCount = versionReader.GetInt64(5);
            sheetChangeCount = versionReader.GetInt64(6);
            summary = versionReader.IsDBNull(7) ? null : versionReader.GetString(7);
        }
        await versionReader.DisposeAsync().ConfigureAwait(false);

        if (options.ExpectedSequence is > 0 && latestSequence is null)
            failures.Add("Expected a captured version, but no version exists.");
        if (latestSequence is not null && latestSequence != currentSequence)
            failures.Add($"Latest version sequence {latestSequence} does not match workbook sequence {currentSequence}.");
        if (options.RequireReadyReport && !string.Equals(reportStatus, "Ready", StringComparison.Ordinal))
            failures.Add($"Expected latest report status Ready, found {reportStatus ?? "<none>"}.");
        if (options.RequireReadyReport && (string.IsNullOrWhiteSpace(reportPath) || !File.Exists(reportPath)))
            failures.Add($"Latest Markdown report is missing: {reportPath ?? "<none>"}.");
        if (options.ExpectedCellChangeCount is { } expectedCellChangeCount && cellChangeCount != expectedCellChangeCount)
            failures.Add($"Expected {expectedCellChangeCount} cell changes, found {cellChangeCount?.ToString() ?? "<none>"}.");
        if (options.ExpectedSheetChangeCount is { } expectedSheetChangeCount && sheetChangeCount != expectedSheetChangeCount)
            failures.Add($"Expected {expectedSheetChangeCount} sheet changes, found {sheetChangeCount?.ToString() ?? "<none>"}.");

        CellChangeResult? cellChange = null;
        if (versionId is not null && options.Address is not null)
        {
            await using var cellCommand = connection.CreateCommand();
            cellCommand.CommandText = """
                SELECT sheet_name,address,kinds,before_json,after_json
                FROM cell_changes WHERE version_id=$version AND address=$address COLLATE NOCASE
                ORDER BY id LIMIT 1;
                """;
            cellCommand.Parameters.AddWithValue("$version", versionId.Value);
            cellCommand.Parameters.AddWithValue("$address", options.Address);
            await using var reader = await cellCommand.ExecuteReaderAsync().ConfigureAwait(false);
            if (await reader.ReadAsync().ConfigureAwait(false))
            {
                cellChange = new CellChangeResult(
                    reader.GetString(0), reader.GetString(1), reader.GetString(2),
                    reader.IsDBNull(3) ? null : reader.GetString(3),
                    reader.IsDBNull(4) ? null : reader.GetString(4));
            }
            else
            {
                failures.Add($"Latest version has no cell change for {options.Address}.");
            }
        }

        if (cellChange is not null)
        {
            if (options.ExpectedKind is not null &&
                !cellChange.Kinds.Split(',').Contains(options.ExpectedKind, StringComparer.Ordinal))
                failures.Add($"Expected change kind {options.ExpectedKind}; found {cellChange.Kinds}.");

            if (options.ExpectedValue is not null)
            {
                var actual = ReadLiteralValue(cellChange.AfterJson);
                if (!string.Equals(actual, options.ExpectedValue, StringComparison.Ordinal))
                    failures.Add($"Expected new literal value '{options.ExpectedValue}', found '{actual ?? "<null>"}'.");
            }
            if (options.ExpectCleared && ReadLiteralValue(cellChange.AfterJson) is not null)
                failures.Add("Expected the latest change to clear the literal value.");
            if (options.ExpectBeforeMissing && cellChange.BeforeJson is not null)
                failures.Add("Expected the changed cell to be absent from the previous snapshot.");
            if (options.ExpectedBeforeValue is not null)
            {
                var actual = ReadLiteralValue(cellChange.BeforeJson);
                if (!string.Equals(actual, options.ExpectedBeforeValue, StringComparison.Ordinal))
                    failures.Add($"Expected old literal value '{options.ExpectedBeforeValue}', found '{actual ?? "<null>"}'.");
            }
            if (options.ExpectedFormulaText is not null)
            {
                var actual = ReadStateString(cellChange.AfterJson, "formulaText");
                if (!string.Equals(actual, options.ExpectedFormulaText, StringComparison.Ordinal))
                    failures.Add($"Expected stored formula text '{options.ExpectedFormulaText}', found '{actual ?? "<null>"}'.");
            }
            if (options.ExpectFormulaMissing && ReadStateString(cellChange.AfterJson, "formulaText") is not null)
                failures.Add("Expected the latest cell state to contain no formula text.");
            if (options.ExpectedCachedResult is not null)
            {
                var actual = ReadStateString(cellChange.AfterJson, "cachedResult");
                if (!string.Equals(actual, options.ExpectedCachedResult, StringComparison.Ordinal))
                    failures.Add($"Expected cached formula result '{options.ExpectedCachedResult}', found '{actual ?? "<null>"}'.");
            }
        }

        SheetChangeResult? sheetChange = null;
        if (versionId is not null && options.ExpectedSheetKind is not null)
        {
            await using var sheetCommand = connection.CreateCommand();
            sheetCommand.CommandText = """
                SELECT sheet_id,kind,before_json,after_json
                FROM sheet_changes WHERE version_id=$version AND kind=$kind
                ORDER BY id LIMIT 1;
                """;
            sheetCommand.Parameters.AddWithValue("$version", versionId.Value);
            sheetCommand.Parameters.AddWithValue("$kind", options.ExpectedSheetKind);
            await using var reader = await sheetCommand.ExecuteReaderAsync().ConfigureAwait(false);
            if (await reader.ReadAsync().ConfigureAwait(false))
            {
                sheetChange = new SheetChangeResult(
                    reader.GetInt64(0), reader.GetString(1),
                    reader.IsDBNull(2) ? null : reader.GetString(2),
                    reader.IsDBNull(3) ? null : reader.GetString(3));
            }
            else
            {
                failures.Add($"Latest version has no sheet change of kind {options.ExpectedSheetKind}.");
            }
        }
        if (sheetChange is not null && options.ExpectedSheetName is not null)
        {
            var beforeName = ReadStateString(sheetChange.BeforeJson, "name");
            var afterName = ReadStateString(sheetChange.AfterJson, "name");
            if (!string.Equals(beforeName, options.ExpectedSheetName, StringComparison.Ordinal) &&
                !string.Equals(afterName, options.ExpectedSheetName, StringComparison.Ordinal))
                failures.Add($"Expected sheet change for '{options.ExpectedSheetName}', found before='{beforeName ?? "<null>"}' after='{afterName ?? "<null>"}'.");
        }

        if (options.ReportContains is not null && reportPath is not null && File.Exists(reportPath))
        {
            var markdown = await File.ReadAllTextAsync(reportPath).ConfigureAwait(false);
            if (!markdown.Contains(options.ReportContains, StringComparison.Ordinal))
                failures.Add($"Markdown report does not contain required text: {options.ReportContains}");
        }

        return new ProbeResult(
            failures.Count == 0,
            failures,
            status,
            currentSequence,
            currentHash,
            lastError,
            errorCount,
            versionCount,
            distinctVersionHashCount,
            reportDirectory,
            versionId is null ? null : new VersionResult(versionId.Value, latestSequence!.Value, versionHash!, reportStatus!, reportPath!, cellChangeCount!.Value, sheetChangeCount!.Value, summary),
            cellChange,
            sheetChange);
    }

    private static string? ReadLiteralValue(string? json)
    {
        if (string.IsNullOrWhiteSpace(json)) return null;
        using var document = JsonDocument.Parse(json);
        return document.RootElement.TryGetProperty("literalValue", out var value) && value.ValueKind != JsonValueKind.Null
            ? value.GetString()
            : null;
    }

    private static string? ReadStateString(string? json, string propertyName)
    {
        if (string.IsNullOrWhiteSpace(json)) return null;
        using var document = JsonDocument.Parse(json);
        return document.RootElement.TryGetProperty(propertyName, out var value) && value.ValueKind != JsonValueKind.Null
            ? value.ToString()
            : null;
    }

    private sealed record Options(
        string DatabasePath,
        string WorkbookPath,
        long? ExpectedSequence,
        string? Address,
        string? ExpectedValue,
        bool ExpectCleared,
        string? ExpectedKind,
        string? ReportContains,
        bool RequireActive,
        bool RequireNoErrors,
        bool RequireNoLastError,
        bool RequireReadyReport,
        string? ExpectedStatus,
        long? MinimumErrors,
        long? ExpectedErrorCount,
        long? ExpectedVersionCount,
        bool RequireUniqueVersionHashes,
        long? ExpectedCellChangeCount,
        long? ExpectedSheetChangeCount,
        bool RequireSourceHashMatch,
        string? ExpectedBeforeValue,
        bool ExpectBeforeMissing,
        string? ExpectedFormulaText,
        bool ExpectFormulaMissing,
        string? ExpectedCachedResult,
        string? ExpectedSheetKind,
        string? ExpectedSheetName)
    {
        public static Options Parse(string[] args)
        {
            var values = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            var switches = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            for (var index = 0; index < args.Length; index++)
            {
                if (!args[index].StartsWith("--", StringComparison.Ordinal))
                    throw new ArgumentException($"Unexpected argument: {args[index]}");
                var key = args[index][2..];
                if (index + 1 < args.Length && !args[index + 1].StartsWith("--", StringComparison.Ordinal))
                    values[key] = args[++index];
                else
                    switches.Add(key);
            }

            if (!values.TryGetValue("database", out var database) || string.IsNullOrWhiteSpace(database))
                throw new ArgumentException("--database is required.");
            if (!values.TryGetValue("workbook", out var workbook) || string.IsNullOrWhiteSpace(workbook))
                throw new ArgumentException("--workbook is required.");
            long? sequence = values.TryGetValue("expected-sequence", out var sequenceText)
                ? long.Parse(sequenceText, System.Globalization.CultureInfo.InvariantCulture)
                : null;
            values.TryGetValue("address", out var address);
            values.TryGetValue("expected-value", out var expectedValue);
            values.TryGetValue("expected-kind", out var expectedKind);
            values.TryGetValue("report-contains", out var reportContains);
            values.TryGetValue("expected-status", out var expectedStatus);
            values.TryGetValue("expected-before-value", out var expectedBeforeValue);
            values.TryGetValue("expected-formula-text", out var expectedFormulaText);
            values.TryGetValue("expected-cached-result", out var expectedCachedResult);
            values.TryGetValue("expected-sheet-kind", out var expectedSheetKind);
            values.TryGetValue("expected-sheet-name", out var expectedSheetName);
            long? minimumErrors = values.TryGetValue("minimum-errors", out var minimumErrorsText)
                ? long.Parse(minimumErrorsText, System.Globalization.CultureInfo.InvariantCulture)
                : null;
            long? expectedErrorCount = values.TryGetValue("expected-error-count", out var expectedErrorCountText)
                ? long.Parse(expectedErrorCountText, System.Globalization.CultureInfo.InvariantCulture)
                : null;
            long? expectedVersionCount = values.TryGetValue("expected-version-count", out var expectedVersionCountText)
                ? long.Parse(expectedVersionCountText, System.Globalization.CultureInfo.InvariantCulture)
                : null;
            long? expectedCellChangeCount = values.TryGetValue("expected-cell-change-count", out var expectedCellChangeCountText)
                ? long.Parse(expectedCellChangeCountText, System.Globalization.CultureInfo.InvariantCulture)
                : null;
            long? expectedSheetChangeCount = values.TryGetValue("expected-sheet-change-count", out var expectedSheetChangeCountText)
                ? long.Parse(expectedSheetChangeCountText, System.Globalization.CultureInfo.InvariantCulture)
                : null;
            return new Options(
                database, workbook, sequence, address, expectedValue,
                switches.Contains("expect-cleared"), expectedKind, reportContains,
                switches.Contains("require-active"), switches.Contains("require-no-errors"),
                switches.Contains("require-no-last-error"), switches.Contains("require-ready-report"),
                expectedStatus, minimumErrors, expectedErrorCount, expectedVersionCount,
                switches.Contains("require-unique-version-hashes"), expectedCellChangeCount,
                expectedSheetChangeCount, switches.Contains("require-source-hash-match"),
                expectedBeforeValue, switches.Contains("expect-before-missing"), expectedFormulaText,
                switches.Contains("expect-formula-missing"), expectedCachedResult, expectedSheetKind,
                expectedSheetName);
        }
    }

    private sealed record ProbeResult(
        bool Passed,
        IReadOnlyList<string> Failures,
        string? WorkbookStatus,
        long? CurrentSequence,
        string? CurrentHash,
        string? LastError,
        long? ErrorCount,
        long? VersionCount,
        long? DistinctVersionHashCount,
        string? ReportDirectory,
        VersionResult? LatestVersion,
        CellChangeResult? CellChange,
        SheetChangeResult? SheetChange);

    private sealed record VersionResult(
        long Id,
        long Sequence,
        string Sha256,
        string ReportStatus,
        string ReportPath,
        long CellChangeCount,
        long SheetChangeCount,
        string? Summary);

    private sealed record CellChangeResult(string SheetName, string Address, string Kinds, string? BeforeJson, string? AfterJson);
    private sealed record SheetChangeResult(long SheetId, string Kind, string? BeforeJson, string? AfterJson);
}
