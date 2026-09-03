using System.Collections.ObjectModel;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ExcelDiffTracker.Core;

// The extraction shape is adapted from xlsx-review (MIT), but narrowed to the
// cell and sheet semantics required by Excel Scenario Analysis Tool.
public sealed class WorkbookExtractor
{
    public WorkbookSnapshot Extract(string path)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(path);
        if (!File.Exists(path))
            throw new FileNotFoundException("Workbook not found.", path);

        WorkbookPackageGuard.Validate(path);

        try
        {
            return ExtractValidated(path);
        }
        catch (UnsafeWorkbookException)
        {
            throw;
        }
        catch (Exception exception) when (exception is OpenXmlPackageException or InvalidDataException or System.Xml.XmlException or KeyNotFoundException or FormatException)
        {
            throw new UnsafeWorkbookException($"The workbook package is corrupt or unreadable: {exception.Message}");
        }
    }

    private static WorkbookSnapshot ExtractValidated(string path)
    {
        using var document = SpreadsheetDocument.Open(path, false);
        var workbookPart = document.WorkbookPart
            ?? throw new UnsafeWorkbookException("The workbook package has no workbook part.");
        var sharedStrings = LoadSharedStrings(workbookPart.SharedStringTablePart?.SharedStringTable);
        var sheets = workbookPart.Workbook.Sheets?.Elements<Sheet>().ToList() ?? [];
        var extractedSheets = new Dictionary<uint, SheetState>();

        for (var position = 0; position < sheets.Count; position++)
        {
            var sheet = sheets[position];
            var sheetId = sheet.SheetId?.Value
                ?? throw new UnsafeWorkbookException("A worksheet is missing its stable sheet identifier.");
            if (extractedSheets.ContainsKey(sheetId))
                throw new UnsafeWorkbookException($"The workbook contains duplicate sheet identifier {sheetId}.");
            var sheetName = sheet.Name?.Value;
            if (string.IsNullOrWhiteSpace(sheetName))
                throw new UnsafeWorkbookException($"Sheet {sheetId} has no name.");
            var cells = new Dictionary<string, CellState>(StringComparer.OrdinalIgnoreCase);

            var relatedPart = GetRelatedPart(workbookPart, sheet);
            if (relatedPart is WorksheetPart worksheetPart)
            {
                foreach (var cell in worksheetPart.Worksheet.Descendants<Cell>())
                {
                    var address = cell.CellReference?.Value;
                    if (!IsValidCellReference(address))
                        throw new UnsafeWorkbookException($"Sheet '{sheetName}' contains an invalid or missing cell reference.");

                    var state = ExtractCell(sheetId, sheetName, address!, cell, sharedStrings);
                    if (state is not null && !cells.TryAdd(address!, state))
                        throw new UnsafeWorkbookException($"Sheet '{sheetName}' contains duplicate cell reference {address}.");
                }
            }
            else if (!IsSupportedNonWorksheetPart(relatedPart))
            {
                throw new UnsafeWorkbookException($"Sheet '{sheetName}' points to an unexpected package part.");
            }

            extractedSheets[sheetId] = new SheetState
            {
                SheetId = sheetId,
                Name = sheetName,
                Position = position,
                Visibility = sheet.State?.Value.ToString() ?? "Visible",
                Cells = new ReadOnlyDictionary<string, CellState>(cells)
            };
        }

        return new WorkbookSnapshot
        {
            SourcePath = Path.GetFullPath(path),
            CapturedAtUtc = DateTime.UtcNow,
            Sheets = new ReadOnlyDictionary<uint, SheetState>(extractedSheets)
        };
    }

    private static bool IsSupportedNonWorksheetPart(OpenXmlPart part) => part.RelationshipType is
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chartsheet" or
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/dialogsheet" or
        "http://schemas.microsoft.com/office/2006/relationships/xlMacrosheet" or
        "http://schemas.microsoft.com/office/2006/relationships/xlIntlMacrosheet";

    private static OpenXmlPart GetRelatedPart(WorkbookPart workbookPart, Sheet sheet)
    {
        var relationshipId = sheet.Id?.Value;
        if (string.IsNullOrWhiteSpace(relationshipId))
            throw new UnsafeWorkbookException("A sheet is missing its package relationship.");
        try
        {
            return workbookPart.GetPartById(relationshipId);
        }
        catch (ArgumentOutOfRangeException exception)
        {
            throw new UnsafeWorkbookException($"A sheet points to a missing package relationship: {exception.Message}");
        }
    }

    private static CellState? ExtractCell(
        uint sheetId,
        string sheetName,
        string address,
        Cell cell,
        IReadOnlyList<string> sharedStrings)
    {
        var formula = cell.CellFormula;
        var (hasValue, value, cellType) = ExtractValue(cell, sharedStrings);
        var hasFormula = formula is not null;

        // A style-only empty cell is deliberately absent from the semantic snapshot.
        if (!hasValue && !hasFormula)
            return null;

        return new CellState
        {
            SheetId = sheetId,
            SheetName = sheetName,
            Address = address.ToUpperInvariant(),
            CellType = hasFormula ? $"formula:{cellType}" : cellType,
            LiteralValue = hasFormula ? null : value,
            FormulaText = formula?.Text,
            FormulaType = formula?.FormulaType?.Value.ToString(),
            FormulaReference = formula?.Reference?.Value,
            FormulaSharedIndex = formula?.SharedIndex?.Value,
            FormulaXml = formula?.OuterXml,
            CachedResult = hasFormula && hasValue ? value : null
        };
    }

    private static (bool HasValue, string? Value, string CellType) ExtractValue(
        Cell cell,
        IReadOnlyList<string> sharedStrings)
    {
        var dataType = cell.DataType?.Value;
        var exactType = dataType?.ToString() ?? "Number";

        if (dataType == CellValues.InlineString)
            return (cell.InlineString is not null, cell.InlineString?.InnerText ?? string.Empty, exactType);

        if (cell.CellValue is null)
            return (false, null, exactType);

        var raw = cell.CellValue.Text;
        if (dataType == CellValues.SharedString)
            return (true, ResolveSharedString(raw, sharedStrings), exactType);
        if (dataType == CellValues.Boolean)
            return (true, NormalizeBoolean(raw), exactType);
        return (true, raw, exactType);
    }

    private static IReadOnlyList<string> LoadSharedStrings(SharedStringTable? table) =>
        table?.Elements<SharedStringItem>().Select(item => item.InnerText).ToArray() ?? [];

    private static string ResolveSharedString(string raw, IReadOnlyList<string> sharedStrings)
    {
        if (!int.TryParse(raw, System.Globalization.NumberStyles.None, System.Globalization.CultureInfo.InvariantCulture, out var index) ||
            index < 0 || index >= sharedStrings.Count)
        {
            throw new UnsafeWorkbookException($"The workbook contains an invalid shared-string reference: {raw}");
        }
        return sharedStrings[index];
    }

    private static string NormalizeBoolean(string raw) => raw switch
    {
        "1" or "true" or "TRUE" => "TRUE",
        "0" or "false" or "FALSE" => "FALSE",
        _ => throw new UnsafeWorkbookException($"The workbook contains an invalid Boolean cell value: {raw}")
    };

    private static bool IsValidCellReference(string? address)
    {
        if (string.IsNullOrWhiteSpace(address) || address.Length is < 2 or > 10)
            return false;

        var index = 0;
        var column = 0;
        while (index < address.Length && char.IsAsciiLetter(address[index]))
        {
            column = column * 26 + char.ToUpperInvariant(address[index]) - 'A' + 1;
            index++;
        }
        if (index is 0 or > 3 || column is < 1 or > 16_384 || index == address.Length)
            return false;

        if (address[index] == '0')
            return false;
        var row = 0;
        while (index < address.Length && char.IsAsciiDigit(address[index]))
        {
            row = row * 10 + address[index] - '0';
            if (row > 1_048_576)
                return false;
            index++;
        }
        return index == address.Length && row > 0;
    }
}
