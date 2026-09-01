using System.Globalization;
using System.Text;
using ExcelDiffTracker.Core;

namespace ExcelDiffTracker.Reporting;

public sealed record ReportContext
{
    public required string WorkbookPath { get; init; }
    public required long Sequence { get; init; }
    public required DateTime CapturedUtc { get; init; }
    public required DateTime FileLastWriteUtc { get; init; }
    public required string PreviousSha256 { get; init; }
    public required string CurrentSha256 { get; init; }
}

public sealed record ReportWriteResult
{
    public required string Path { get; init; }
    public required bool WasTruncated { get; init; }
    public required int IncludedCellChanges { get; init; }
}

public sealed class MarkdownReportWriter
{
    public const int DefaultInlineChangeLimit = 5_000;

    public async Task<ReportWriteResult> WriteAsync(
        string outputPath,
        ReportContext context,
        WorkbookDiff diff,
        int inlineChangeLimit = DefaultInlineChangeLimit,
        bool includeAll = false,
        CancellationToken cancellationToken = default)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(outputPath);
        ArgumentNullException.ThrowIfNull(context);
        ArgumentNullException.ThrowIfNull(diff);

        var effectiveLimit = includeAll ? int.MaxValue : Math.Max(0, inlineChangeLimit);
        var includedCells = Math.Min(diff.CellChanges.Count, effectiveLimit);
        var wasTruncated = includedCells < diff.CellChanges.Count;
        var markdown = Render(context, diff, includedCells, wasTruncated);

        var fullPath = Path.GetFullPath(outputPath);
        var directory = Path.GetDirectoryName(fullPath)
            ?? throw new InvalidOperationException("The report path has no parent directory.");
        Directory.CreateDirectory(directory);

        var temporaryPath = Path.Combine(directory, $".{Path.GetFileName(fullPath)}.{Guid.NewGuid():N}.tmp");
        try
        {
            await File.WriteAllTextAsync(temporaryPath, markdown, new UTF8Encoding(false), cancellationToken).ConfigureAwait(false);
            File.Move(temporaryPath, fullPath, overwrite: true);
        }
        finally
        {
            if (File.Exists(temporaryPath))
                File.Delete(temporaryPath);
        }

        return new ReportWriteResult
        {
            Path = fullPath,
            WasTruncated = wasTruncated,
            IncludedCellChanges = includedCells
        };
    }

    public string Render(ReportContext context, WorkbookDiff diff, int includedCellChanges, bool wasTruncated)
    {
        var builder = new StringBuilder(16_384);
        builder.AppendLine("# Excel Diff Tracker report");
        builder.AppendLine();
        builder.AppendLine($"- Workbook: <code>{EscapeUntrusted(context.WorkbookPath)}</code>");
        builder.AppendLine($"- Version: {context.Sequence.ToString(CultureInfo.InvariantCulture)}");
        builder.AppendLine($"- Captured (UTC): {FormatDate(context.CapturedUtc)}");
        builder.AppendLine($"- Workbook saved (UTC): {FormatDate(context.FileLastWriteUtc)}");
        builder.AppendLine($"- Previous SHA-256: `{context.PreviousSha256}`");
        builder.AppendLine($"- Current SHA-256: `{context.CurrentSha256}`");
        builder.AppendLine();

        AppendSummary(builder, diff);
        AppendSheetChanges(builder, diff.SheetChanges);
        AppendCellChanges(builder, diff.CellChanges.Take(includedCellChanges), diff.CellChanges.Count, wasTruncated);
        return builder.ToString();
    }

    private static void AppendSummary(StringBuilder builder, WorkbookDiff diff)
    {
        builder.AppendLine("## Summary");
        builder.AppendLine();
        if (!diff.HasTrackedChanges)
        {
            builder.AppendLine("No tracked changes. The workbook file changed, but only ignored content or package metadata differed.");
            builder.AppendLine();
            return;
        }

        builder.AppendLine("| Category | Count |");
        builder.AppendLine("|---|---:|");
        builder.AppendLine($"| Sheet changes | {diff.SheetChanges.Count.ToString("N0", CultureInfo.InvariantCulture)} |");
        builder.AppendLine($"| Changed cells | {diff.CellChanges.Count.ToString("N0", CultureInfo.InvariantCulture)} |");
        builder.AppendLine($"| Literal value changes | {diff.LiteralChangeCount.ToString("N0", CultureInfo.InvariantCulture)} |");
        builder.AppendLine($"| Formula text changes | {diff.FormulaChangeCount.ToString("N0", CultureInfo.InvariantCulture)} |");
        builder.AppendLine($"| Calculated result changes | {diff.FormulaResultChangeCount.ToString("N0", CultureInfo.InvariantCulture)} |");
        builder.AppendLine($"| Cell type changes | {diff.CellChanges.Count(change => change.Kinds.Contains(CellChangeKind.CellTypeChanged)).ToString("N0", CultureInfo.InvariantCulture)} |");
        builder.AppendLine();
    }

    private static void AppendSheetChanges(StringBuilder builder, IReadOnlyList<SheetDelta> changes)
    {
        builder.AppendLine("## Sheet changes");
        builder.AppendLine();
        if (changes.Count == 0)
        {
            builder.AppendLine("None.");
            builder.AppendLine();
            return;
        }

        builder.AppendLine("| Change | Before | After |");
        builder.AppendLine("|---|---|---|");
        foreach (var change in changes)
        {
            builder.Append("| ").Append(Humanize(change.Kind)).Append(" | ")
                .Append(DescribeSheet(change.Before)).Append(" | ")
                .Append(DescribeSheet(change.After)).AppendLine(" |");
        }
        builder.AppendLine();
    }

    private static void AppendCellChanges(
        StringBuilder builder,
        IEnumerable<CellDelta> changes,
        int totalCount,
        bool wasTruncated)
    {
        builder.AppendLine("## Cell changes");
        builder.AppendLine();
        if (totalCount == 0)
        {
            builder.AppendLine("None.");
            builder.AppendLine();
            return;
        }

        if (wasTruncated)
        {
            builder.AppendLine($"> This automatic report includes the first {changes.Count():N0} of {totalCount:N0} cell changes. The complete history remains in the local database; use **Export full Markdown** in the app for an untruncated report.");
            builder.AppendLine();
        }

        foreach (var change in changes)
        {
            builder.Append("### ").Append(EscapeHeading(change.SheetName)).Append('!').AppendLine(change.Address);
            builder.AppendLine();
            builder.AppendLine($"**Changes:** {string.Join(", ", change.Kinds.Select(Humanize))}");
            builder.AppendLine();
            builder.AppendLine("| Field | Before | After |");
            builder.AppendLine("|---|---|---|");
            AppendCellRow(builder, "Type", change.Before?.CellType, change.After?.CellType);
            AppendCellRow(builder, "Literal value", change.Before?.LiteralValue, change.After?.LiteralValue);
            AppendCellRow(builder, "Formula", change.Before?.FormulaText, change.After?.FormulaText);
            AppendCellRow(builder, "Formula type", change.Before?.FormulaType, change.After?.FormulaType);
            AppendCellRow(builder, "Formula reference", change.Before?.FormulaReference, change.After?.FormulaReference);
            AppendCellRow(builder, "Shared formula index", change.Before?.FormulaSharedIndex?.ToString(CultureInfo.InvariantCulture), change.After?.FormulaSharedIndex?.ToString(CultureInfo.InvariantCulture));
            AppendCellRow(builder, "Stored formula XML", change.Before?.FormulaXml, change.After?.FormulaXml);
            AppendCellRow(builder, "Calculated result", change.Before?.CachedResult, change.After?.CachedResult);
            builder.AppendLine();
        }
    }

    private static void AppendCellRow(StringBuilder builder, string field, string? before, string? after) =>
        builder.Append("| ").Append(field).Append(" | ").Append(DisplayValue(before)).Append(" | ").Append(DisplayValue(after)).AppendLine(" |");

    private static string DescribeSheet(SheetState? sheet) => sheet is null
        ? "—"
        : $"{EscapeTable(sheet.Name)} (position {sheet.Position + 1}, {EscapeTable(sheet.Visibility)})";

    private static string DisplayValue(string? value)
    {
        if (value is null)
            return "`<none>`";
        if (value.Length == 0)
            return "`<empty string>`";
        return $"<pre>{EscapeUntrusted(value)}</pre>";
    }

    private static string Humanize<T>(T value) where T : Enum
    {
        var text = value.ToString();
        var builder = new StringBuilder(text.Length + 8);
        for (var index = 0; index < text.Length; index++)
        {
            if (index > 0 && char.IsUpper(text[index]))
                builder.Append(' ');
            builder.Append(index == 0 ? char.ToUpperInvariant(text[index]) : char.ToLowerInvariant(text[index]));
        }
        return builder.ToString();
    }

    private static string EscapeTable(string value) => EscapeUntrusted(value).Replace("|", "&#124;", StringComparison.Ordinal);
    private static string EscapeHeading(string value) => EscapeUntrusted(value);
    private static string EscapeUntrusted(string value) => EscapeHtml(value)
        .Replace("\\", "&#92;", StringComparison.Ordinal)
        .Replace("`", "&#96;", StringComparison.Ordinal)
        .Replace("*", "&#42;", StringComparison.Ordinal)
        .Replace("_", "&#95;", StringComparison.Ordinal)
        .Replace("[", "&#91;", StringComparison.Ordinal)
        .Replace("]", "&#93;", StringComparison.Ordinal)
        .Replace("(", "&#40;", StringComparison.Ordinal)
        .Replace(")", "&#41;", StringComparison.Ordinal)
        .Replace("!", "&#33;", StringComparison.Ordinal)
        .Replace(":", "&#58;", StringComparison.Ordinal)
        .Replace("/", "&#47;", StringComparison.Ordinal)
        .Replace("@", "&#64;", StringComparison.Ordinal);
    private static string EscapeHtml(string value) => value
        .Replace("&", "&amp;", StringComparison.Ordinal)
        .Replace("<", "&lt;", StringComparison.Ordinal)
        .Replace(">", "&gt;", StringComparison.Ordinal)
        .Replace("\r\n", "<br>", StringComparison.Ordinal)
        .Replace("\n", "<br>", StringComparison.Ordinal)
        .Replace("\r", "<br>", StringComparison.Ordinal)
        .Replace("|", "&#124;", StringComparison.Ordinal);
    private static string FormatDate(DateTime value) => value.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss 'UTC'", CultureInfo.InvariantCulture);
}
