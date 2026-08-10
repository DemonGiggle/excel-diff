using ExcelDiff.Models;
using ExcelDiff.Services;
using Xunit;

namespace ExcelDiff.Tests;

public sealed class WorksheetDiffEngineTests
{
    private readonly WorksheetDiffEngine _engine = new();

    [Fact]
    public void InsertedRow_IsAlignedWithoutChangingFollowingRows()
    {
        var oldSheet = Grid(["Header"], ["Alice"], ["Bob"]);
        var newSheet = Grid(["Header"], ["Inserted"], ["Alice"], ["Bob"]);

        var result = Compare(oldSheet, newSheet);

        Assert.Equal(1, result.DifferenceRowCount);
        var added = Assert.Single(result.Rows, row => row.Kind == UnifiedRowKind.Added);
        Assert.Equal(2, added.NewRowNumber);
        Assert.Equal("Inserted", added.Cells[0].NewValue);
        Assert.Equal(UnifiedCellKind.Added, added.Cells[0].Kind);
    }

    [Fact]
    public void RemovedRow_IsAlignedWithoutChangingFollowingRows()
    {
        var oldSheet = Grid(["Header"], ["Removed"], ["Alice"], ["Bob"]);
        var newSheet = Grid(["Header"], ["Alice"], ["Bob"]);

        var result = Compare(oldSheet, newSheet);

        var removed = Assert.Single(result.Rows, row => row.Kind == UnifiedRowKind.Removed);
        Assert.Equal("Removed", removed.Cells[0].OldValue);
        Assert.Equal(UnifiedCellKind.Removed, removed.Cells[0].Kind);
    }

    [Fact]
    public void InsertedColumn_IsShownAsAddedCells()
    {
        var oldSheet = Grid(["Name", "City"], ["Alice", "Paris"]);
        var newSheet = Grid(["Name", "Team", "City"], ["Alice", "Blue", "Paris"]);

        var result = Compare(oldSheet, newSheet);

        Assert.Equal(3, result.Columns.Count);
        Assert.Null(result.Columns[1].OldIndex);
        Assert.Equal(1, result.Columns[1].NewIndex);
        Assert.All(result.Rows.Where(row => row.IsDifference), row => Assert.Equal(UnifiedCellKind.Added, row.Cells[1].Kind));
    }

    [Fact]
    public void ChangedCell_DisplaysOldAndNewValues()
    {
        var oldSheet = Grid(["ID", "Name"], ["1", "Alice"], ["2", "Bob"]);
        var newSheet = Grid(["ID", "Name"], ["1", "Alicia"], ["2", "Bob"]);

        var result = Compare(oldSheet, newSheet);

        var changed = Assert.Single(result.Rows, row => row.Kind == UnifiedRowKind.Changed);
        Assert.Equal(UnifiedCellKind.Changed, changed.Cells[1].Kind);
        Assert.Equal("Alice", changed.Cells[1].OldValue);
        Assert.Equal("Alicia", changed.Cells[1].NewValue);
    }

    [Fact]
    public void ConsecutiveUnchangedRows_AreFoldedAroundChangedRow()
    {
        var oldSheet = Grid(["Header"], ["1"], ["2"], ["old"], ["4"], ["5"]);
        var newSheet = Grid(["Header"], ["1"], ["2"], ["new"], ["4"], ["5"]);

        var result = Compare(oldSheet, newSheet);

        Assert.Collection(result.Rows,
            row => { Assert.Equal(UnifiedRowKind.Fold, row.Kind); Assert.Equal(2, row.FoldedRowCount); },
            row => Assert.Equal(UnifiedRowKind.Changed, row.Kind),
            row => { Assert.Equal(UnifiedRowKind.Fold, row.Kind); Assert.Equal(2, row.FoldedRowCount); });
    }

    [Fact]
    public void IdenticalFirstRow_BecomesHeadersAndIsRemovedFromBody()
    {
        var oldSheet = Grid(["Employee ID", "Name"], ["1", "Alice"]);
        var newSheet = Grid(["Employee ID", "Name"], ["1", "Alicia"]);

        var result = Compare(oldSheet, newSheet);

        Assert.True(result.FirstRowUsedAsHeaders);
        Assert.Equal(["Employee ID", "Name"], result.Columns.Select(column => column.Label));
        var row = Assert.Single(result.Rows);
        Assert.Equal(2, row.OldRowNumber);
        Assert.Equal(2, row.NewRowNumber);
    }

    [Fact]
    public void ChangedFirstRow_RemainsInBodyAndKeepsColumnLetters()
    {
        var oldSheet = Grid(["Employee ID", "Old name"], ["1", "Alice"]);
        var newSheet = Grid(["Employee ID", "New name"], ["1", "Alice"]);

        var result = Compare(oldSheet, newSheet);

        Assert.False(result.FirstRowUsedAsHeaders);
        Assert.Equal(["A", "B"], result.Columns.Select(column => column.Label));
        Assert.Contains(result.Rows, row => row.OldRowNumber == 1 && row.NewRowNumber == 1 && row.Kind == UnifiedRowKind.Changed);
    }

    [Fact]
    public void EmptyInteriorCells_RemainPartOfUsedRectangle()
    {
        var oldSheet = Grid(["A", null, "C"], [null, null, "last"]);
        var newSheet = Grid(["A", "added", "C"], [null, null, "last"]);

        var result = Compare(oldSheet, newSheet);

        Assert.Equal(3, result.Columns.Count);
        Assert.Contains(result.Rows, row => row.Cells[1].Kind == UnifiedCellKind.Added && row.Cells[1].NewValue == "added");
    }

    private WorksheetDiffResult Compare(WorksheetGrid oldSheet, WorksheetGrid newSheet) =>
        _engine.Compare(oldSheet, newSheet, CancellationToken.None);

    private static WorksheetGrid Grid(params string?[][] values)
    {
        var rows = new Dictionary<int, GridRow>();
        var maxColumn = 0;
        for (var rowIndex = 0; rowIndex < values.Length; rowIndex++)
        {
            var cells = new Dictionary<int, CellData>();
            for (var column = 0; column < values[rowIndex].Length; column++)
            {
                if (values[rowIndex][column] is not string value) continue;
                cells[column] = new CellData(CellValueKind.Text, value, value);
                maxColumn = Math.Max(maxColumn, column + 1);
            }
            if (cells.Count > 0) rows[rowIndex + 1] = new GridRow(rowIndex + 1, cells);
        }
        return new WorksheetGrid("test.xlsx", "Sheet1", values.Length, maxColumn, rows, []);
    }
}
