namespace ExcelDiff.Models;

public sealed record GridRow(int RowNumber, IReadOnlyDictionary<int, CellData> Cells)
{
    public CellData CellAt(int zeroBasedColumn) => Cells.GetValueOrDefault(zeroBasedColumn, CellData.Blank);
}

public sealed record WorksheetGrid(
    string FilePath,
    string SheetName,
    int MaxRow,
    int MaxColumn,
    IReadOnlyDictionary<int, GridRow> Rows,
    IReadOnlyList<ComparisonIssue> ReadIssues)
{
    public CellData CellAt(int oneBasedRow, int zeroBasedColumn) =>
        Rows.TryGetValue(oneBasedRow, out var row) ? row.CellAt(zeroBasedColumn) : CellData.Blank;
}

public enum UnifiedCellKind { Unchanged, Changed, Added, Removed }
public enum UnifiedRowKind { Unchanged, Changed, Added, Removed, Fold }

public sealed record UnifiedColumn(int? OldIndex, int? NewIndex, string Label);

public sealed record UnifiedCell(UnifiedCellKind Kind, string OldValue, string NewValue)
{
    public string AccessibleText => Kind switch
    {
        UnifiedCellKind.Changed => $"Changed from {OldValue} to {NewValue}",
        UnifiedCellKind.Added => $"Added {NewValue}",
        UnifiedCellKind.Removed => $"Removed {OldValue}",
        _ => NewValue
    };
}

public sealed class UnifiedDiffRow
{
    public required UnifiedRowKind Kind { get; init; }
    public required int? OldRowNumber { get; init; }
    public required int? NewRowNumber { get; init; }
    public required IReadOnlyList<UnifiedCell> Cells { get; init; }
    public int FoldedRowCount { get; init; }
    public bool IsDifference => Kind is UnifiedRowKind.Changed or UnifiedRowKind.Added or UnifiedRowKind.Removed;
    public string RowLabel => Kind switch
    {
        UnifiedRowKind.Fold => $"{FoldedRowCount:N0} unchanged row{(FoldedRowCount == 1 ? "" : "s")}",
        UnifiedRowKind.Added => $"+{NewRowNumber}",
        UnifiedRowKind.Removed => $"−{OldRowNumber}",
        _ when OldRowNumber != NewRowNumber => $"{OldRowNumber} → {NewRowNumber}",
        _ => (NewRowNumber ?? OldRowNumber)?.ToString() ?? string.Empty
    };
}

public sealed record WorksheetDiffResult(
    IReadOnlyList<UnifiedColumn> Columns,
    IReadOnlyList<UnifiedDiffRow> Rows,
    bool FirstRowUsedAsHeaders,
    int DifferenceRowCount,
    int AddedRowCount,
    int RemovedRowCount,
    int ChangedCellCount,
    IReadOnlyList<ComparisonIssue> Issues);
