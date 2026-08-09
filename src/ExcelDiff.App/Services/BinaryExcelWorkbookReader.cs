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

    public WorksheetData ReadSheet(
        string filePath,
        string sheetName,
        int headerRow,
        CancellationToken cancellationToken,
        IProgress<int>? progress = null)
    {
        ValidatePath(filePath);
        if (headerRow < 1) throw new ArgumentOutOfRangeException(nameof(headerRow), "Header row must be 1 or greater.");

        using var reader = OpenReader(filePath);
        do
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (string.Equals(reader.Name, sheetName, StringComparison.Ordinal))
                return ReadCurrentSheet(reader, filePath, sheetName, headerRow, cancellationToken, progress);
        } while (reader.NextResult());

        throw new InvalidDataException($"Sheet '{sheetName}' was not found.");
    }

    private static WorksheetData ReadCurrentSheet(
        IExcelDataReader reader,
        string filePath,
        string sheetName,
        int headerRow,
        CancellationToken cancellationToken,
        IProgress<int>? progress)
    {
        var issues = new List<ComparisonIssue>
        {
            new(IssueSeverity.Information, "Legacy workbook",
                "Legacy .xls formulas are compared using their saved results; formulas are never executed.", sheetName)
        };
        var rows = new List<SheetRow>();
        List<SheetHeader>? headers = null;
        var rowNumber = 0;
        var estimatedRows = TryGetRowCount(reader);

        while (reader.Read())
        {
            cancellationToken.ThrowIfCancellationRequested();
            rowNumber++;
            if (rowNumber < headerRow) continue;

            if (rowNumber == headerRow)
            {
                headers = BuildHeaders(reader, issues, sheetName);
                continue;
            }

            if (headers is null) continue;
            var values = new CellData[headers.Count];
            var hasAnyValue = false;
            for (var column = 0; column < headers.Count; column++)
            {
                var value = column < reader.FieldCount ? ReadCell(reader, column) : CellData.Blank;
                values[column] = value;
                hasAnyValue |= value.Kind != CellValueKind.Blank;
            }

            if (hasAnyValue) rows.Add(new SheetRow(rowNumber, values));
            if (estimatedRows > 0 && rowNumber % 250 == 0)
                progress?.Report(Math.Clamp(rowNumber * 100 / estimatedRows, 0, 99));
        }

        if (headers is null)
            throw new InvalidDataException($"Header row {headerRow} was not found in '{sheetName}'.");
        if (rows.Count == 0)
            issues.Add(new ComparisonIssue(IssueSeverity.Warning, "Empty sheet", "No data rows were found below the selected header row.", sheetName));
        progress?.Report(100);
        return new WorksheetData(filePath, sheetName, headerRow, headers, rows, issues);
    }

    private static List<SheetHeader> BuildHeaders(IExcelDataReader reader, ICollection<ComparisonIssue> issues, string sheetName)
    {
        var lastColumn = -1;
        for (var column = 0; column < reader.FieldCount; column++)
        {
            if (ReadCell(reader, column).Kind != CellValueKind.Blank) lastColumn = column;
        }
        if (lastColumn < 0) throw new InvalidDataException($"The selected header row in '{sheetName}' is empty.");

        var headers = new List<SheetHeader>(lastColumn + 1);
        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        for (var column = 0; column <= lastColumn; column++)
        {
            var text = ReadCell(reader, column).DisplayValue.Trim();
            if (string.IsNullOrWhiteSpace(text))
            {
                text = $"Unnamed column {GetColumnName(column)}";
                issues.Add(new ComparisonIssue(IssueSeverity.Warning, "Header", "A blank header was assigned a temporary name.", $"{sheetName}!{GetColumnName(column)}"));
            }
            if (!seen.Add(text))
            {
                issues.Add(new ComparisonIssue(IssueSeverity.Warning, "Header", $"Duplicate header '{text}' was disambiguated.", $"{sheetName}!{GetColumnName(column)}"));
                text = $"{text} [{GetColumnName(column)}]";
            }
            headers.Add(new SheetHeader(column, text));
        }
        return headers;
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

    private static string GetColumnName(int zeroBasedIndex)
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
