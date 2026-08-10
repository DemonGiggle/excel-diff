using System.Globalization;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using ExcelDataReader;
using ExcelDiff.Models;

namespace ExcelDiff.Services;

public sealed partial class BinaryExcelWorkbookReader : IWorkbookReader
{
    static BinaryExcelWorkbookReader() => Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

    public IReadOnlyList<string> GetSheetNames(string filePath)
    {
        ValidatePath(filePath);
        using var reader = OpenReader(filePath);
        var names = new List<string>();
        do
        {
            if (!string.IsNullOrWhiteSpace(reader.Name)) names.Add(reader.Name);
        } while (reader.NextResult());
        return names;
    }

    public WorksheetGrid ReadGrid(
        string filePath,
        string sheetName,
        CancellationToken cancellationToken,
        IProgress<int>? progress = null)
    {
        ValidatePath(filePath);
        using var reader = OpenReader(filePath);
        do
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (!string.Equals(reader.Name, sheetName, StringComparison.Ordinal)) continue;

            var issues = new List<ComparisonIssue>
            {
                new(IssueSeverity.Information, "Legacy workbook", "Legacy .xls formulas are compared using saved values; formulas are never executed.", sheetName)
            };
            var rows = new Dictionary<int, GridRow>();
            var rowNumber = 0;
            var maxRow = 0;
            var maxColumn = 0;
            var estimatedRows = TryGetRowCount(reader);
            while (reader.Read())
            {
                cancellationToken.ThrowIfCancellationRequested();
                rowNumber++;
                var cells = new Dictionary<int, CellData>();
                for (var column = 0; column < reader.FieldCount; column++)
                {
                    var value = ReadCell(reader, column);
                    if (value.Kind == CellValueKind.Blank) continue;
                    cells[column] = value;
                    maxColumn = Math.Max(maxColumn, column + 1);
                }
                if (cells.Count > 0)
                {
                    rows[rowNumber] = new GridRow(rowNumber, cells);
                    maxRow = rowNumber;
                }
                if (estimatedRows > 0 && rowNumber % 250 == 0)
                    progress?.Report(Math.Clamp(rowNumber * 100 / estimatedRows, 0, 99));
            }
            var mergedRanges = reader.MergeCells.Select(range => new MergedCellRange(
                range.FromRow + 1,
                range.FromColumn,
                range.ToRow + 1,
                range.ToColumn)).ToArray();
            MergedCellNormalizer.Expand(rows, mergedRanges, ref maxRow, ref maxColumn);
            if (maxRow == 0 || maxColumn == 0)
                issues.Add(new ComparisonIssue(IssueSeverity.Warning, "Empty sheet", "The selected sheet contains no saved cell values.", sheetName));
            progress?.Report(100);
            return new WorksheetGrid(filePath, sheetName, maxRow, maxColumn, rows, issues)
            {
                MergedRanges = mergedRanges
            };
        } while (reader.NextResult());

        throw new InvalidDataException($"Sheet '{sheetName}' was not found.");
    }

    private static CellData ReadCell(IExcelDataReader reader, int column)
    {
        var value = reader.GetValue(column);
        if (value is null or DBNull) return CellData.Blank;
        if (value is string text) return new CellData(CellValueKind.Text, text, text);
        if (value is bool boolean)
            return new CellData(CellValueKind.Boolean, boolean ? "True" : "False", boolean ? "true" : "false");
        if (value is DateTime date) return Date(date);
        if (value is TimeSpan time)
        {
            var canonicalTime = time.ToString("c", CultureInfo.InvariantCulture);
            return new CellData(CellValueKind.Text, canonicalTime, canonicalTime);
        }
        if (IsNumeric(value))
        {
            var number = Convert.ToDouble(value, CultureInfo.InvariantCulture);
            var format = reader.GetNumberFormatString(column);
            if (IsDateFormat(format) && number is >= -657435.0 and <= 2958465.99999999)
                return Date(DateTime.FromOADate(number));
            var canonical = CanonicalNumber(value);
            return new CellData(CellValueKind.Number, canonical, canonical);
        }
        if (value.GetType().Name.Contains("ExcelError", StringComparison.OrdinalIgnoreCase))
        {
            var error = Convert.ToString(value, CultureInfo.InvariantCulture) ?? "Excel error";
            return new CellData(CellValueKind.Error, error, error);
        }

        var display = Convert.ToString(value, CultureInfo.InvariantCulture) ?? string.Empty;
        return new CellData(CellValueKind.Text, display, display);
    }

    private static CellData Date(DateTime date)
    {
        var canonical = date.ToString("O", CultureInfo.InvariantCulture);
        var display = date.TimeOfDay == TimeSpan.Zero
            ? date.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture)
            : date.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture);
        return new CellData(CellValueKind.Date, display, canonical);
    }

    private static string CanonicalNumber(object value)
    {
        try
        {
            return Convert.ToDecimal(value, CultureInfo.InvariantCulture).ToString("G29", CultureInfo.InvariantCulture);
        }
        catch (OverflowException)
        {
            return Convert.ToDouble(value, CultureInfo.InvariantCulture).ToString("R", CultureInfo.InvariantCulture);
        }
    }

    private static bool IsNumeric(object value) => value is byte or sbyte or short or ushort or int or uint or long or ulong or float or double or decimal;

    private static bool IsDateFormat(string? format) =>
        !string.IsNullOrWhiteSpace(format) && DateFormatTokenRegex().IsMatch(QuotedFormatRegex().Replace(format, string.Empty));

    [GeneratedRegex("[ymdhis]", RegexOptions.IgnoreCase)] private static partial Regex DateFormatTokenRegex();
    [GeneratedRegex("\"[^\"]*\"")] private static partial Regex QuotedFormatRegex();

    private static int TryGetRowCount(IExcelDataReader reader)
    {
        try { return reader.RowCount; }
        catch (InvalidOperationException) { return 0; }
    }

    private static IExcelDataReader OpenReader(string path)
    {
        var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        try
        {
            return ExcelReaderFactory.CreateBinaryReader(stream, new ExcelReaderConfiguration
            {
                LeaveOpen = false,
                FallbackEncoding = Encoding.GetEncoding(1252)
            });
        }
        catch
        {
            stream.Dispose();
            throw;
        }
    }

    private static void ValidatePath(string filePath)
    {
        if (!File.Exists(filePath)) throw new FileNotFoundException("The selected workbook could not be found.", filePath);
        if (!string.Equals(Path.GetExtension(filePath), ".xls", StringComparison.OrdinalIgnoreCase))
            throw new NotSupportedException("This reader supports legacy .xls workbooks only.");
    }
}
