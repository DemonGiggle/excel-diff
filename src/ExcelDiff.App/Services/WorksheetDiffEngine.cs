using System.Text;
using ExcelDiff.Models;

namespace ExcelDiff.Services;

public interface IWorksheetDiffEngine
{
    WorksheetDiffResult Compare(WorksheetGrid oldSheet, WorksheetGrid newSheet, CancellationToken cancellationToken, IProgress<int>? progress = null);
}

public sealed class WorksheetDiffEngine : IWorksheetDiffEngine
{
    public WorksheetDiffResult Compare(
        WorksheetGrid oldSheet,
        WorksheetGrid newSheet,
        CancellationToken cancellationToken,
        IProgress<int>? progress = null)
    {
        var initialRows = Align(
            Enumerable.Range(1, oldSheet.MaxRow).Select(row => RawRowSignature(oldSheet, row)).ToArray(),
            Enumerable.Range(1, newSheet.MaxRow).Select(row => RawRowSignature(newSheet, row)).ToArray());
        progress?.Report(15);

        var columns = Align(
            Enumerable.Range(0, oldSheet.MaxColumn).Select(column => ColumnSignature(oldSheet, column, initialRows, true)).ToArray(),
            Enumerable.Range(0, newSheet.MaxColumn).Select(column => ColumnSignature(newSheet, column, initialRows, false)).ToArray());
        progress?.Report(30);

        var rows = Align(
            Enumerable.Range(1, oldSheet.MaxRow).Select(row => AlignedRowSignature(oldSheet, row, columns, true)).ToArray(),
            Enumerable.Range(1, newSheet.MaxRow).Select(row => AlignedRowSignature(newSheet, row, columns, false)).ToArray());
        progress?.Report(45);

        // A second column pass uses the aligned rows, allowing row and column insertions
        // in the same workbook to reinforce each other.
        columns = Align(
            Enumerable.Range(0, oldSheet.MaxColumn).Select(column => ColumnSignature(oldSheet, column, rows, true)).ToArray(),
            Enumerable.Range(0, newSheet.MaxColumn).Select(column => ColumnSignature(newSheet, column, rows, false)).ToArray());

        var unifiedColumns = columns.Select(pair => new UnifiedColumn(
            pair.OldIndex,
            pair.NewIndex,
            ColumnLabel(pair.OldIndex, pair.NewIndex))).ToArray();
        var rawRows = new List<UnifiedDiffRow>(rows.Count);
        var changedCells = 0;

        for (var index = 0; index < rows.Count; index++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var pair = rows[index];
            var oldRow = pair.OldIndex is int oldIndex ? oldIndex + 1 : (int?)null;
            var newRow = pair.NewIndex is int newIndex ? newIndex + 1 : (int?)null;
            var cells = new UnifiedCell[columns.Count];
            for (var columnIndex = 0; columnIndex < columns.Count; columnIndex++)
            {
                var column = columns[columnIndex];
                var oldCell = oldRow is int oldRowNumber && column.OldIndex is int oldColumn
                    ? oldSheet.CellAt(oldRowNumber, oldColumn)
                    : CellData.Blank;
                var newCell = newRow is int newRowNumber && column.NewIndex is int newColumn
                    ? newSheet.CellAt(newRowNumber, newColumn)
                    : CellData.Blank;
                cells[columnIndex] = CompareCell(oldCell, newCell);
                if (cells[columnIndex].Kind != UnifiedCellKind.Unchanged) changedCells++;
            }

            var kind = oldRow is null
                ? UnifiedRowKind.Added
                : newRow is null
                    ? UnifiedRowKind.Removed
                    : cells.Any(cell => cell.Kind != UnifiedCellKind.Unchanged)
                        ? UnifiedRowKind.Changed
                        : UnifiedRowKind.Unchanged;
            rawRows.Add(new UnifiedDiffRow
            {
                Kind = kind,
                OldRowNumber = oldRow,
                NewRowNumber = newRow,
                Cells = cells
            });
            if (index % 100 == 0)
                progress?.Report(45 + (rows.Count == 0 ? 45 : index * 45 / rows.Count));
        }

        var foldedRows = FoldUnchangedRows(rawRows, unifiedColumns.Length);
        progress?.Report(100);
        return new WorksheetDiffResult(
            unifiedColumns,
            foldedRows,
            rawRows.Count(row => row.IsDifference),
            rawRows.Count(row => row.Kind == UnifiedRowKind.Added),
            rawRows.Count(row => row.Kind == UnifiedRowKind.Removed),
            changedCells,
            oldSheet.ReadIssues.Concat(newSheet.ReadIssues).ToArray());
    }

    private static UnifiedCell CompareCell(CellData oldCell, CellData newCell)
    {
        if (Equivalent(oldCell, newCell))
            return new UnifiedCell(UnifiedCellKind.Unchanged, oldCell.DisplayValue, newCell.DisplayValue);
        if (oldCell.Kind == CellValueKind.Blank)
            return new UnifiedCell(UnifiedCellKind.Added, string.Empty, newCell.DisplayValue);
        if (newCell.Kind == CellValueKind.Blank)
            return new UnifiedCell(UnifiedCellKind.Removed, oldCell.DisplayValue, string.Empty);
        return new UnifiedCell(UnifiedCellKind.Changed, oldCell.DisplayValue, newCell.DisplayValue);
    }

    private static bool Equivalent(CellData oldCell, CellData newCell) =>
        oldCell.Kind == newCell.Kind && string.Equals(oldCell.CanonicalValue, newCell.CanonicalValue, StringComparison.Ordinal);

    private static IReadOnlyList<UnifiedDiffRow> FoldUnchangedRows(IReadOnlyList<UnifiedDiffRow> rows, int columnCount)
    {
        var result = new List<UnifiedDiffRow>();
        for (var index = 0; index < rows.Count;)
        {
            if (rows[index].Kind != UnifiedRowKind.Unchanged)
            {
                result.Add(rows[index++]);
                continue;
            }

            var start = index;
            while (index < rows.Count && rows[index].Kind == UnifiedRowKind.Unchanged) index++;
            var count = index - start;
            if (count == 1)
            {
                result.Add(rows[start]);
                continue;
            }

            result.Add(new UnifiedDiffRow
            {
                Kind = UnifiedRowKind.Fold,
                OldRowNumber = rows[start].OldRowNumber,
                NewRowNumber = rows[start].NewRowNumber,
                FoldedRowCount = count,
                Cells = Enumerable.Repeat(new UnifiedCell(UnifiedCellKind.Unchanged, string.Empty, string.Empty), columnCount).ToArray()
            });
        }
        return result;
    }

    private static string RawRowSignature(WorksheetGrid sheet, int row)
    {
        var builder = new StringBuilder();
        for (var column = 0; column < sheet.MaxColumn; column++) Append(builder, sheet.CellAt(row, column));
        return builder.ToString();
    }

    private static string AlignedRowSignature(WorksheetGrid sheet, int row, IReadOnlyList<Alignment> columns, bool old)
    {
        var builder = new StringBuilder();
        foreach (var pair in columns)
        {
            var column = old ? pair.OldIndex : pair.NewIndex;
            if (column is int value) Append(builder, sheet.CellAt(row, value));
        }
        return builder.ToString();
    }

    private static string ColumnSignature(WorksheetGrid sheet, int column, IReadOnlyList<Alignment> rows, bool old)
    {
        var builder = new StringBuilder();
        foreach (var pair in rows)
        {
            var rowIndex = old ? pair.OldIndex : pair.NewIndex;
            if (rowIndex is int value) Append(builder, sheet.CellAt(value + 1, column));
        }
        return builder.ToString();
    }

    private static void Append(StringBuilder builder, CellData cell) =>
        builder.Append((int)cell.Kind).Append(':').Append(cell.CanonicalValue.Length).Append(':').Append(cell.CanonicalValue).Append('|');

    private static List<Alignment> Align(IReadOnlyList<string> oldValues, IReadOnlyList<string> newValues)
    {
        var oldPositions = Positions(oldValues);
        var newPositions = Positions(newValues);
        var candidates = oldPositions
            .Where(entry => entry.Value.Count == 1 && newPositions.TryGetValue(entry.Key, out var positions) && positions.Count == 1)
            .Select(entry => new Alignment(entry.Value[0], newPositions[entry.Key][0]))
            .OrderBy(pair => pair.OldIndex)
            .ToArray();
        var anchors = LongestIncreasingSubsequence(candidates);
        var result = new List<Alignment>();
        var oldStart = 0;
        var newStart = 0;
        foreach (var anchor in anchors.Append(new Alignment(oldValues.Count, newValues.Count)))
        {
            AddGap(result, oldStart, anchor.OldIndex!.Value, newStart, anchor.NewIndex!.Value);
            if (anchor.OldIndex < oldValues.Count && anchor.NewIndex < newValues.Count) result.Add(anchor);
            oldStart = anchor.OldIndex.Value + 1;
            newStart = anchor.NewIndex.Value + 1;
        }
        return result;
    }

    private static Dictionary<string, List<int>> Positions(IReadOnlyList<string> values)
    {
        var result = new Dictionary<string, List<int>>(StringComparer.Ordinal);
        for (var index = 0; index < values.Count; index++)
        {
            if (!result.TryGetValue(values[index], out var positions)) result[values[index]] = positions = [];
            positions.Add(index);
        }
        return result;
    }

    private static IReadOnlyList<Alignment> LongestIncreasingSubsequence(IReadOnlyList<Alignment> candidates)
    {
        if (candidates.Count == 0) return [];
        var tails = new int[candidates.Count];
        var previous = Enumerable.Repeat(-1, candidates.Count).ToArray();
        var length = 0;
        for (var index = 0; index < candidates.Count; index++)
        {
            var low = 0;
            var high = length;
            while (low < high)
            {
                var middle = (low + high) / 2;
                if (candidates[tails[middle]].NewIndex < candidates[index].NewIndex) low = middle + 1;
                else high = middle;
            }
            if (low > 0) previous[index] = tails[low - 1];
            tails[low] = index;
            if (low == length) length++;
        }
        var result = new List<Alignment>();
        for (var index = tails[length - 1]; index >= 0; index = previous[index])
        {
            result.Add(candidates[index]);
            if (previous[index] < 0) break;
        }
        result.Reverse();
        return result;
    }

    private static void AddGap(List<Alignment> result, int oldStart, int oldEnd, int newStart, int newEnd)
    {
        var oldCount = Math.Max(0, oldEnd - oldStart);
        var newCount = Math.Max(0, newEnd - newStart);
        var common = Math.Min(oldCount, newCount);
        for (var index = 0; index < common; index++) result.Add(new Alignment(oldStart + index, newStart + index));
        for (var index = common; index < oldCount; index++) result.Add(new Alignment(oldStart + index, null));
        for (var index = common; index < newCount; index++) result.Add(new Alignment(null, newStart + index));
    }

    private static string ColumnLabel(int? oldIndex, int? newIndex)
    {
        var index = newIndex ?? oldIndex ?? 0;
        var prefix = oldIndex is null ? "+ " : newIndex is null ? "− " : string.Empty;
        return prefix + ExcelColumnName(index);
    }

    private static string ExcelColumnName(int zeroBasedIndex)
    {
        var value = zeroBasedIndex + 1;
        var name = string.Empty;
        while (value > 0)
        {
            value--;
            name = (char)('A' + value % 26) + name;
            value /= 26;
        }
        return name;
    }

    private sealed record Alignment(int? OldIndex, int? NewIndex);
}
