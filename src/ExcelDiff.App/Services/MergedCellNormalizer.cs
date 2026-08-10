using ExcelDiff.Models;

namespace ExcelDiff.Services;

internal sealed record MergedCellRange(int FromRow, int FromColumn, int ToRow, int ToColumn);

internal static class MergedCellNormalizer
{
    public static void Expand(
        Dictionary<int, GridRow> rows,
        IEnumerable<MergedCellRange> ranges,
        ref int maxRow,
        ref int maxColumn)
    {
        foreach (var range in ranges)
        {
            if (range.FromRow < 1 || range.FromColumn < 0 || range.ToRow < range.FromRow || range.ToColumn < range.FromColumn)
                continue;
            var value = rows.TryGetValue(range.FromRow, out var sourceRow)
                ? sourceRow.CellAt(range.FromColumn)
                : CellData.Blank;
            // An empty merged region is formatting only. The comparison remains value-based,
            // so it does not enlarge the used rectangle.
            if (value.Kind == CellValueKind.Blank) continue;

            for (var rowNumber = range.FromRow; rowNumber <= range.ToRow; rowNumber++)
            {
                var cells = rows.TryGetValue(rowNumber, out var existingRow)
                    ? new Dictionary<int, CellData>(existingRow.Cells)
                    : [];
                for (var column = range.FromColumn; column <= range.ToColumn; column++)
                    cells[column] = value;
                rows[rowNumber] = new GridRow(rowNumber, cells);
            }
            maxRow = Math.Max(maxRow, range.ToRow);
            maxColumn = Math.Max(maxColumn, range.ToColumn + 1);
        }
    }
}
