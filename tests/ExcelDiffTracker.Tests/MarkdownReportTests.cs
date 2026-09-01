using ExcelDiffTracker.Core;
using ExcelDiffTracker.Reporting;

namespace ExcelDiffTracker.Tests;

public sealed class MarkdownReportTests
{
    [Fact]
    public async Task AutomaticReportTruncatesAfterFiveThousandButFullExportDoesNot()
    {
        using var directory = new TestDirectory();
        var changes = Enumerable.Range(1, 5_001).Select(index => new CellDelta
        {
            SheetId = 1,
            SheetName = "Data|Sheet",
            Address = $"A{index}",
            Kinds = [CellChangeKind.LiteralChanged],
            Before = Cell($"old|{index}\nline"),
            After = Cell($"new<{index}>")
        }).ToList();
        var diff = new WorkbookDiff { CellChanges = changes, SheetChanges = [] };
        var context = new ReportContext
        {
            WorkbookPath = @"C:\Data\Book.xlsx",
            Sequence = 2,
            CapturedUtc = DateTime.UtcNow,
            FileLastWriteUtc = DateTime.UtcNow,
            PreviousSha256 = new string('a', 64),
            CurrentSha256 = new string('b', 64)
        };
        var writer = new MarkdownReportWriter();
        var automaticPath = Path.Combine(directory.Path, "automatic.md");
        var fullPath = Path.Combine(directory.Path, "full.md");

        var automatic = await writer.WriteAsync(automaticPath, context, diff);
        var full = await writer.WriteAsync(fullPath, context, diff, includeAll: true);

        Assert.True(automatic.WasTruncated);
        Assert.Equal(5_000, automatic.IncludedCellChanges);
        Assert.False(full.WasTruncated);
        Assert.Equal(5_001, full.IncludedCellChanges);
        Assert.Contains("first 5,000 of 5,001", await File.ReadAllTextAsync(automaticPath));
        Assert.Contains("### Data&#124;Sheet!A5001", await File.ReadAllTextAsync(fullPath));
    }

    [Fact]
    public void WorkbookControlledTextCannotCreateMarkdownLinksOrHtml()
    {
        var cell = new CellState
        {
            SheetId = 1,
            SheetName = "![sheet](https://example.invalid)",
            Address = "A1",
            CellType = "str",
            LiteralValue = "<img src=https://example.invalid/track> [open](https://example.invalid)"
        };
        var diff = new WorkbookDiff
        {
            SheetChanges = [],
            CellChanges =
            [
                new CellDelta
                {
                    SheetId = 1,
                    SheetName = cell.SheetName,
                    Address = cell.Address,
                    Kinds = [CellChangeKind.LiteralAdded],
                    After = cell
                }
            ]
        };
        var context = new ReportContext
        {
            WorkbookPath = @"C:\[private](https://example.invalid)\book.xlsx",
            Sequence = 1,
            CapturedUtc = DateTime.UtcNow,
            FileLastWriteUtc = DateTime.UtcNow,
            PreviousSha256 = new string('a', 64),
            CurrentSha256 = new string('b', 64)
        };

        var markdown = new MarkdownReportWriter().Render(context, diff, 1, false);

        Assert.DoesNotContain("https://", markdown, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("<img", markdown, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("&#91;sheet&#93;&#40;https&#58;&#47;&#47;example.invalid&#41;", markdown);
    }

    private static CellState Cell(string value) => new()
    {
        SheetId = 1,
        SheetName = "Data|Sheet",
        Address = "A1",
        CellType = "str",
        LiteralValue = value
    };
}
