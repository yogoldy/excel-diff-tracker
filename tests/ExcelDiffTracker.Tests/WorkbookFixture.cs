using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ExcelDiffTracker.Tests;

internal sealed record FixtureCell(string Address, string? Value = null, string? Formula = null, CellValues? Type = null, uint StyleIndex = 0);
internal sealed record FixtureSheet(uint Id, string Name, SheetStateValues? State, params FixtureCell[] Cells);

internal static class WorkbookFixture
{
    public static void Create(string path, bool macroEnabled, params FixtureSheet[] sheets)
    {
        Directory.CreateDirectory(Path.GetDirectoryName(path)!);
        using var document = SpreadsheetDocument.Create(
            path,
            macroEnabled ? SpreadsheetDocumentType.MacroEnabledWorkbook : SpreadsheetDocumentType.Workbook);
        var workbookPart = document.AddWorkbookPart();
        workbookPart.Workbook = new Workbook();
        if (macroEnabled)
        {
            var vbaProject = workbookPart.AddNewPart<VbaProjectPart>();
            using var macroBytes = new MemoryStream([0x01, 0x00, 0x00, 0x00, 0x56, 0x42, 0x41]);
            vbaProject.FeedData(macroBytes);
        }
        var styles = workbookPart.AddNewPart<WorkbookStylesPart>();
        styles.Stylesheet = new Stylesheet(
            new Fonts(new Font()),
            new Fills(new Fill(new PatternFill { PatternType = PatternValues.None })),
            new Borders(new Border()),
            new CellStyleFormats(new CellFormat()),
            new CellFormats(new CellFormat(), new CellFormat { FontId = 0, FillId = 0, BorderId = 0 }));
        styles.Stylesheet.Save();

        var sheetList = workbookPart.Workbook.AppendChild(new Sheets());
        foreach (var fixtureSheet in sheets)
        {
            var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
            var sheetData = new SheetData();
            foreach (var rowGroup in fixtureSheet.Cells.GroupBy(cell => Row(cell.Address)).OrderBy(group => group.Key))
            {
                var row = new Row { RowIndex = rowGroup.Key };
                foreach (var fixtureCell in rowGroup)
                {
                    var cell = new Cell
                    {
                        CellReference = fixtureCell.Address,
                        StyleIndex = fixtureCell.StyleIndex
                    };
                    if (fixtureCell.Type is { } type)
                        cell.DataType = type;
                    if (fixtureCell.Formula is not null)
                        cell.CellFormula = new CellFormula(fixtureCell.Formula);
                    if (fixtureCell.Value is not null)
                        cell.CellValue = new CellValue(fixtureCell.Value);
                    row.Append(cell);
                }
                sheetData.Append(row);
            }
            worksheetPart.Worksheet = new Worksheet(sheetData);
            worksheetPart.Worksheet.Save();
            sheetList.Append(new Sheet
            {
                Id = workbookPart.GetIdOfPart(worksheetPart),
                SheetId = fixtureSheet.Id,
                Name = fixtureSheet.Name,
                State = fixtureSheet.State
            });
        }
        workbookPart.Workbook.Save();
    }

    private static uint Row(string address) => uint.Parse(new string(address.SkipWhile(char.IsAsciiLetter).ToArray()), System.Globalization.CultureInfo.InvariantCulture);
}
