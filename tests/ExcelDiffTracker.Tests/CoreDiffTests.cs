using System.Collections.ObjectModel;
using DocumentFormat.OpenXml.Spreadsheet;
using ExcelDiffTracker.Core;

namespace ExcelDiffTracker.Tests;

public sealed class CoreDiffTests
{
    [Fact]
    public void ExtractsValuesFormulasTypesAndCachedResults()
    {
        using var directory = new TestDirectory();
        var path = Path.Combine(directory.Path, "sample.xlsx");
        WorkbookFixture.Create(path, false, new FixtureSheet(1, "Data", null,
            new FixtureCell("A1", "hello", Type: CellValues.String),
            new FixtureCell("B1", "1", Type: CellValues.Boolean),
            new FixtureCell("C1", "3", "1+2")));

        var snapshot = new WorkbookExtractor().Extract(path);
        var cells = snapshot.Sheets[1].Cells;
        Assert.Equal("hello", cells["A1"].LiteralValue);
        Assert.Equal("TRUE", cells["B1"].LiteralValue);
        Assert.Equal("1+2", cells["C1"].FormulaText);
        Assert.Equal("3", cells["C1"].CachedResult);
        Assert.NotNull(cells["C1"].FormulaXml);
    }

    [Fact]
    public void SeparatesFormulaTextAndCalculatedResultChanges()
    {
        var before = Snapshot(Cell(formula: "A2+1", result: "2"));
        var resultOnly = Snapshot(Cell(formula: "A2+1", result: "3"));
        var formulaAndResult = Snapshot(Cell(formula: "A2+2", result: "4"));
        var differ = new WorkbookDiffer();

        var resultDelta = Assert.Single(differ.Compare(before, resultOnly).CellChanges);
        Assert.Equal([CellChangeKind.FormulaResultChanged], resultDelta.Kinds);
        var formulaDelta = Assert.Single(differ.Compare(before, formulaAndResult).CellChanges);
        Assert.Contains(CellChangeKind.FormulaChanged, formulaDelta.Kinds);
        Assert.Contains(CellChangeKind.FormulaResultChanged, formulaDelta.Kinds);
    }

    [Fact]
    public void DetectsSheetRenameReorderAndVisibilityUsingStableId()
    {
        var before = Workbook(
            Sheet(10, "First", 0, "Visible"),
            Sheet(20, "Second", 1, "Visible"));
        var after = Workbook(
            Sheet(20, "Renamed", 0, "Hidden"),
            Sheet(10, "First", 1, "Visible"));
        var changes = new WorkbookDiffer().Compare(before, after).SheetChanges;

        Assert.Contains(changes, item => item.SheetId == 20 && item.Kind == SheetChangeKind.Renamed);
        Assert.Contains(changes, item => item.SheetId == 20 && item.Kind == SheetChangeKind.VisibilityChanged);
        Assert.Equal(2, changes.Count(item => item.Kind == SheetChangeKind.Reordered));
    }

    [Fact]
    public void StyleOnlyWorkbookChangeHasNoSemanticDelta()
    {
        using var directory = new TestDirectory();
        var beforePath = Path.Combine(directory.Path, "before.xlsx");
        var afterPath = Path.Combine(directory.Path, "after.xlsx");
        WorkbookFixture.Create(beforePath, false, new FixtureSheet(1, "Sheet1", null, new FixtureCell("A1", "42", StyleIndex: 0)));
        WorkbookFixture.Create(afterPath, false, new FixtureSheet(1, "Sheet1", null, new FixtureCell("A1", "42", StyleIndex: 1)));
        var extractor = new WorkbookExtractor();

        var diff = new WorkbookDiffer().Compare(extractor.Extract(beforePath), extractor.Extract(afterPath));
        Assert.False(diff.HasTrackedChanges);
    }

    [Fact]
    public void MacroEnabledWorkbookIsReadWithoutExecutingMacros()
    {
        using var directory = new TestDirectory();
        var path = Path.Combine(directory.Path, "macro.xlsm");
        WorkbookFixture.Create(path, true, new FixtureSheet(1, "Sheet1", null, new FixtureCell("A1", "2", "1+1")));
        using (var document = DocumentFormat.OpenXml.Packaging.SpreadsheetDocument.Open(path, false))
            Assert.NotNull(document.WorkbookPart?.VbaProjectPart);
        var hashBefore = Convert.ToHexString(System.Security.Cryptography.SHA256.HashData(File.ReadAllBytes(path)));

        var snapshot = new WorkbookExtractor().Extract(path);

        var hashAfter = Convert.ToHexString(System.Security.Cryptography.SHA256.HashData(File.ReadAllBytes(path)));
        Assert.Equal("1+1", snapshot.Sheets[1].Cells["A1"].FormulaText);
        Assert.Equal(hashBefore, hashAfter);
    }

    [Fact]
    public void DetectsLiteralFormulaAndTypeTransitionMatrix()
    {
        var differ = new WorkbookDiffer();
        var added = Assert.Single(differ.Compare(Workbook(), Snapshot(Cell(literal: "1"))).CellChanges);
        Assert.Contains(CellChangeKind.LiteralAdded, added.Kinds);

        var edited = Assert.Single(differ.Compare(Snapshot(Cell(literal: "1")), Snapshot(Cell(literal: "2"))).CellChanges);
        Assert.Contains(CellChangeKind.LiteralChanged, edited.Kinds);

        var formula = Assert.Single(differ.Compare(Snapshot(Cell(literal: "2")), Snapshot(Cell(formula: "1+1", result: "2"))).CellChanges);
        Assert.Contains(CellChangeKind.FormulaAdded, formula.Kinds);
        Assert.Contains(CellChangeKind.LiteralCleared, formula.Kinds);
        Assert.Contains(CellChangeKind.CellTypeChanged, formula.Kinds);

        var literalAgain = Assert.Single(differ.Compare(Snapshot(Cell(formula: "1+1", result: "2")), Snapshot(Cell(literal: "2"))).CellChanges);
        Assert.Contains(CellChangeKind.FormulaRemoved, literalAgain.Kinds);
        Assert.Contains(CellChangeKind.LiteralAdded, literalAgain.Kinds);

        var cleared = Assert.Single(differ.Compare(Snapshot(Cell(literal: "2")), Workbook()).CellChanges);
        Assert.Contains(CellChangeKind.LiteralCleared, cleared.Kinds);
    }

    [Fact]
    public void DetectsSheetAdditionAndRemovalWithStableIdentifiers()
    {
        var before = Workbook(Sheet(1, "Existing", 0, "Visible"));
        var after = Workbook(Sheet(2, "Replacement", 0, "Visible"));
        var changes = new WorkbookDiffer().Compare(before, after).SheetChanges;

        Assert.Contains(changes, item => item.SheetId == 1 && item.Kind == SheetChangeKind.Removed);
        Assert.Contains(changes, item => item.SheetId == 2 && item.Kind == SheetChangeKind.Added);
    }

    [Theory]
    [InlineData("legacy.xls")]
    [InlineData("binary.xlsb")]
    public void RejectsUnsupportedExcelFormats(string filename)
    {
        using var directory = new TestDirectory();
        var path = Path.Combine(directory.Path, filename);
        File.WriteAllText(path, "unsupported");
        Assert.Throws<UnsupportedWorkbookException>(() => new WorkbookExtractor().Extract(path));
    }

    private static CellState Cell(string? literal = null, string? formula = null, string? result = null) => new()
    {
        SheetId = 1, SheetName = "Sheet1", Address = "A1", CellType = formula is null ? "n" : "formula:n",
        LiteralValue = literal, FormulaText = formula, FormulaXml = formula is null ? null : $"<f>{formula}</f>", CachedResult = result
    };
    private static WorkbookSnapshot Snapshot(CellState cell) => Workbook(new SheetState
    {
        SheetId = 1, Name = "Sheet1", Position = 0, Visibility = "Visible",
        Cells = new ReadOnlyDictionary<string, CellState>(new Dictionary<string, CellState>(StringComparer.OrdinalIgnoreCase) { [cell.Address] = cell })
    });
    private static SheetState Sheet(uint id, string name, int position, string visibility) => new()
    {
        SheetId = id, Name = name, Position = position, Visibility = visibility,
        Cells = new ReadOnlyDictionary<string, CellState>(new Dictionary<string, CellState>())
    };
    private static WorkbookSnapshot Workbook(params SheetState[] sheets) => new()
    {
        SourcePath = "fixture.xlsx", CapturedAtUtc = DateTime.UtcNow,
        Sheets = new ReadOnlyDictionary<uint, SheetState>(sheets.ToDictionary(sheet => sheet.SheetId))
    };
}
