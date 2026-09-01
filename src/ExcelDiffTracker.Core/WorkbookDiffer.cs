namespace ExcelDiffTracker.Core;

// The high-level diff design was informed by the MIT-licensed xlsx-review
// project at commit c39bcf7. This is an independent C# implementation; see
// THIRD-PARTY-NOTICES.md for attribution and license details.

public sealed class WorkbookDiffer
{
    public WorkbookDiff Compare(WorkbookSnapshot before, WorkbookSnapshot after)
    {
        ArgumentNullException.ThrowIfNull(before);
        ArgumentNullException.ThrowIfNull(after);

        var sheetChanges = new List<SheetDelta>();
        var cellChanges = new List<CellDelta>();
        var allSheetIds = before.Sheets.Keys.Union(after.Sheets.Keys).OrderBy(id => id);

        foreach (var sheetId in allSheetIds)
        {
            before.Sheets.TryGetValue(sheetId, out var oldSheet);
            after.Sheets.TryGetValue(sheetId, out var newSheet);

            if (oldSheet is null && newSheet is not null)
            {
                sheetChanges.Add(new SheetDelta { SheetId = sheetId, Kind = SheetChangeKind.Added, After = newSheet });
                AddWholeSheetCellChanges(cellChanges, newSheet, added: true);
                continue;
            }

            if (oldSheet is not null && newSheet is null)
            {
                sheetChanges.Add(new SheetDelta { SheetId = sheetId, Kind = SheetChangeKind.Removed, Before = oldSheet });
                AddWholeSheetCellChanges(cellChanges, oldSheet, added: false);
                continue;
            }

            if (oldSheet is null || newSheet is null)
                continue;

            if (!string.Equals(oldSheet.Name, newSheet.Name, StringComparison.Ordinal))
                sheetChanges.Add(new SheetDelta { SheetId = sheetId, Kind = SheetChangeKind.Renamed, Before = oldSheet, After = newSheet });
            if (oldSheet.Position != newSheet.Position)
                sheetChanges.Add(new SheetDelta { SheetId = sheetId, Kind = SheetChangeKind.Reordered, Before = oldSheet, After = newSheet });
            if (!string.Equals(oldSheet.Visibility, newSheet.Visibility, StringComparison.Ordinal))
                sheetChanges.Add(new SheetDelta { SheetId = sheetId, Kind = SheetChangeKind.VisibilityChanged, Before = oldSheet, After = newSheet });

            CompareCells(oldSheet, newSheet, cellChanges);
        }

        return new WorkbookDiff
        {
            CellChanges = cellChanges
                .OrderBy(change => change.SheetName, StringComparer.OrdinalIgnoreCase)
                .ThenBy(change => AddressSortKey(change.Address))
                .ToList(),
            SheetChanges = sheetChanges
                .OrderBy(change => change.After?.Position ?? change.Before?.Position ?? int.MaxValue)
                .ThenBy(change => change.Kind)
                .ToList()
        };
    }

    private static void AddWholeSheetCellChanges(List<CellDelta> changes, SheetState sheet, bool added)
    {
        foreach (var cell in sheet.Cells.Values)
        {
            var kinds = new List<CellChangeKind>();
            if (cell.HasFormula)
                kinds.Add(added ? CellChangeKind.FormulaAdded : CellChangeKind.FormulaRemoved);
            else
                kinds.Add(added ? CellChangeKind.LiteralAdded : CellChangeKind.LiteralCleared);

            changes.Add(new CellDelta
            {
                SheetId = sheet.SheetId,
                SheetName = sheet.Name,
                Address = cell.Address,
                Kinds = kinds,
                Before = added ? null : cell,
                After = added ? cell : null
            });
        }
    }

    private static void CompareCells(SheetState oldSheet, SheetState newSheet, List<CellDelta> changes)
    {
        var allAddresses = oldSheet.Cells.Keys.Union(newSheet.Cells.Keys, StringComparer.OrdinalIgnoreCase);
        foreach (var address in allAddresses)
        {
            oldSheet.Cells.TryGetValue(address, out var oldCell);
            newSheet.Cells.TryGetValue(address, out var newCell);
            var kinds = DetermineKinds(oldCell, newCell);
            if (kinds.Count == 0)
                continue;

            changes.Add(new CellDelta
            {
                SheetId = newSheet.SheetId,
                SheetName = newSheet.Name,
                Address = address.ToUpperInvariant(),
                Kinds = kinds,
                Before = oldCell,
                After = newCell
            });
        }
    }

    private static IReadOnlyList<CellChangeKind> DetermineKinds(CellState? oldCell, CellState? newCell)
    {
        var kinds = new List<CellChangeKind>();

        if (oldCell is null && newCell is not null)
        {
            kinds.Add(newCell.HasFormula ? CellChangeKind.FormulaAdded : CellChangeKind.LiteralAdded);
            return kinds;
        }

        if (oldCell is not null && newCell is null)
        {
            kinds.Add(oldCell.HasFormula ? CellChangeKind.FormulaRemoved : CellChangeKind.LiteralCleared);
            return kinds;
        }

        if (oldCell is null || newCell is null)
            return kinds;

        if (!string.Equals(oldCell.CellType, newCell.CellType, StringComparison.Ordinal))
            kinds.Add(CellChangeKind.CellTypeChanged);

        if (oldCell.HasFormula || newCell.HasFormula)
        {
            if (!oldCell.HasFormula && newCell.HasFormula)
                kinds.Add(CellChangeKind.FormulaAdded);
            else if (oldCell.HasFormula && !newCell.HasFormula)
                kinds.Add(CellChangeKind.FormulaRemoved);
            else if (!string.Equals(oldCell.FormulaIdentity, newCell.FormulaIdentity, StringComparison.Ordinal))
                kinds.Add(CellChangeKind.FormulaChanged);

            if (!string.Equals(oldCell.CachedResult, newCell.CachedResult, StringComparison.Ordinal))
                kinds.Add(CellChangeKind.FormulaResultChanged);

            if (oldCell.HasFormula && !newCell.HasFormula && newCell.LiteralValue is not null)
                kinds.Add(CellChangeKind.LiteralAdded);
            if (!oldCell.HasFormula && newCell.HasFormula && oldCell.LiteralValue is not null)
                kinds.Add(CellChangeKind.LiteralCleared);
        }
        else if (!string.Equals(oldCell.LiteralValue, newCell.LiteralValue, StringComparison.Ordinal))
        {
            kinds.Add(oldCell.LiteralValue is null
                ? CellChangeKind.LiteralAdded
                : newCell.LiteralValue is null
                    ? CellChangeKind.LiteralCleared
                    : CellChangeKind.LiteralChanged);
        }

        return kinds.Distinct().ToList();
    }

    private static string AddressSortKey(string address)
    {
        var letters = new string(address.TakeWhile(char.IsAsciiLetter).ToArray()).ToUpperInvariant();
        var digits = new string(address.SkipWhile(char.IsAsciiLetter).ToArray());
        var column = 0;
        foreach (var letter in letters)
            column = column * 26 + letter - 'A' + 1;
        _ = int.TryParse(digits, out var row);
        return $"{row:D10}:{column:D7}";
    }
}
