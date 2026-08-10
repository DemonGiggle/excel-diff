using System.Globalization;
using System.IO;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using ExcelDiff.Models;

namespace ExcelDiff.Services;

public sealed partial class OpenXmlWorkbookReader : IWorkbookReader
{
    public IReadOnlyList<string> GetSheetNames(string filePath)
    {
        ValidatePath(filePath);
        using var stream = OpenReadOnly(filePath);
        using var document = SpreadsheetDocument.Open(stream, false);
        var workbookPart = document.WorkbookPart ?? throw new InvalidDataException("The workbook structure is missing.");
        var sheets = workbookPart.Workbook?.Sheets ?? throw new InvalidDataException("The workbook has no worksheet list.");
        return sheets.Elements<Sheet>()
            .Select(s => s.Name?.Value)
            .Where(n => !string.IsNullOrWhiteSpace(n))
            .Cast<string>()
            .ToArray() ?? [];
    }

    public WorksheetGrid ReadGrid(
        string filePath,
        string sheetName,
        CancellationToken cancellationToken,
        IProgress<int>? progress = null)
    {
        ValidatePath(filePath);
        using var stream = OpenReadOnly(filePath);
        using var document = SpreadsheetDocument.Open(stream, false);
        var workbookPart = document.WorkbookPart ?? throw new InvalidDataException("The workbook structure is missing.");
        var sheets = workbookPart.Workbook?.Sheets ?? throw new InvalidDataException("The workbook has no worksheet list.");
        var sheet = sheets.Elements<Sheet>().FirstOrDefault(item => string.Equals(item.Name?.Value, sheetName, StringComparison.Ordinal));
        if (sheet?.Id?.Value is null) throw new InvalidDataException($"Sheet '{sheetName}' was not found.");

        var worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id.Value);
        var sharedStrings = workbookPart.SharedStringTablePart?.SharedStringTable;
        var styles = workbookPart.WorkbookStylesPart?.Stylesheet;
        var issues = new List<ComparisonIssue>();
        var rows = new Dictionary<int, GridRow>();
        var maxRow = 0;
        var maxColumn = 0;
        var estimatedLastRow = ReadEstimatedLastRow(worksheetPart);

        using var reader = OpenXmlReader.Create(worksheetPart);
        while (reader.Read())
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (!reader.IsStartElement || reader.ElementType != typeof(Row) || reader.LoadCurrentElement() is not Row row) continue;
            var rowNumber = checked((int)(row.RowIndex?.Value ?? 0));
            if (rowNumber < 1) continue;
            var cells = new Dictionary<int, CellData>();
            foreach (var cell in row.Elements<Cell>())
            {
                var column = GetColumnIndex(cell.CellReference?.Value);
                var value = ReadCell(cell, sharedStrings, styles);
                if (value.MissingCachedFormulaValue)
                    issues.Add(new ComparisonIssue(IssueSeverity.Warning, "Formula", "A formula has no saved result and is treated as blank.", $"{sheetName}!{GetColumnName(column)}{rowNumber}"));
                if (value.Kind == CellValueKind.Blank) continue;
                cells[column] = value;
                maxColumn = Math.Max(maxColumn, column + 1);
            }
            if (cells.Count > 0)
            {
                rows[rowNumber] = new GridRow(rowNumber, cells);
                maxRow = Math.Max(maxRow, rowNumber);
            }
            if (estimatedLastRow > 0 && rowNumber % 250 == 0)
                progress?.Report(Math.Clamp(rowNumber * 100 / estimatedLastRow, 0, 99));
        }

        if (maxRow == 0 || maxColumn == 0)
            issues.Add(new ComparisonIssue(IssueSeverity.Warning, "Empty sheet", "The selected sheet contains no saved cell values.", sheetName));
        progress?.Report(100);
        return new WorksheetGrid(filePath, sheetName, maxRow, maxColumn, rows, issues);
    }

    private static CellData ReadCell(Cell cell, SharedStringTable? sharedStrings, Stylesheet? styles)
    {
        var hasFormula = cell.CellFormula is not null;
        var raw = cell.CellValue?.Text;
        if (hasFormula && raw is null)
            return new CellData(CellValueKind.Blank, string.Empty, string.Empty, true, true);

        if (cell.DataType?.Value == CellValues.InlineString)
        {
            var text = cell.InlineString?.InnerText ?? string.Empty;
            return Text(text, hasFormula);
        }

        if (raw is null) return CellData.Blank;
        var dataType = cell.DataType?.Value;
        if (dataType == CellValues.SharedString)
        {
            if (int.TryParse(raw, NumberStyles.Integer, CultureInfo.InvariantCulture, out var stringIndex))
                return Text(sharedStrings?.Elements<SharedStringItem>().ElementAtOrDefault(stringIndex)?.InnerText ?? raw, hasFormula);
            return Text(raw, hasFormula);
        }
        if (dataType == CellValues.String) return Text(raw, hasFormula);
        if (dataType == CellValues.Boolean)
        {
            var boolean = raw == "1" || bool.TryParse(raw, out var parsed) && parsed;
            return new CellData(CellValueKind.Boolean, boolean ? "True" : "False", boolean ? "true" : "false", hasFormula);
        }
        if (dataType == CellValues.Error) return new CellData(CellValueKind.Error, raw, raw, hasFormula);
        if (dataType == CellValues.Date)
        {
            if (DateTime.TryParse(raw, CultureInfo.InvariantCulture, DateTimeStyles.RoundtripKind, out var isoDate))
                return Date(isoDate, hasFormula);
            return Text(raw, hasFormula);
        }

        if (decimal.TryParse(raw, NumberStyles.Float, CultureInfo.InvariantCulture, out var number))
        {
            if (IsDateCell(cell, styles) && double.TryParse(raw, NumberStyles.Float, CultureInfo.InvariantCulture, out var oa) && oa is >= -657435.0 and <= 2958465.99999999)
                return Date(DateTime.FromOADate(oa), hasFormula);
            var canonical = number.ToString("G29", CultureInfo.InvariantCulture);
            return new CellData(CellValueKind.Number, canonical, canonical, hasFormula);
        }
        return Text(raw, hasFormula);
    }

    private static CellData Text(string text, bool hasFormula) => new(CellValueKind.Text, text, text, hasFormula);
    private static CellData Date(DateTime date, bool hasFormula)
    {
        var canonical = date.ToString("O", CultureInfo.InvariantCulture);
        var display = date.TimeOfDay == TimeSpan.Zero ? date.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture) : date.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture);
        return new CellData(CellValueKind.Date, display, canonical, hasFormula);
    }

    private static bool IsDateCell(Cell cell, Stylesheet? styles)
    {
        var cellFormats = styles?.CellFormats;
        if (cellFormats is null || cell.StyleIndex?.Value is not uint styleIndex) return false;
        var format = cellFormats.Elements<CellFormat>().ElementAtOrDefault((int)styleIndex);
        if (format?.NumberFormatId?.Value is not uint id) return false;
        if (id is >= 14 and <= 22 or >= 45 and <= 47) return true;
        var code = styles?.NumberingFormats?.Elements<NumberingFormat>()
            .FirstOrDefault(n => n.NumberFormatId?.Value == id)?.FormatCode?.Value;
        return code is not null && DateFormatTokenRegex().IsMatch(StripQuotedSections(code));
    }

    private static string StripQuotedSections(string formatCode) => QuotedFormatRegex().Replace(formatCode, string.Empty);
    [GeneratedRegex("[ymdhis]", RegexOptions.IgnoreCase)] private static partial Regex DateFormatTokenRegex();
    [GeneratedRegex("\"[^\"]*\"")] private static partial Regex QuotedFormatRegex();

    private static int ReadEstimatedLastRow(WorksheetPart part)
    {
        var reference = part.Worksheet?.SheetDimension?.Reference?.Value;
        if (string.IsNullOrWhiteSpace(reference)) return 0;
        var end = reference.Split(':').Last();
        var digits = new string(end.Where(char.IsDigit).ToArray());
        return int.TryParse(digits, out var row) ? row : 0;
    }

    private static int GetColumnIndex(string? reference)
    {
        if (string.IsNullOrWhiteSpace(reference)) return 0;
        var value = 0;
        foreach (var c in reference)
        {
            if (!char.IsLetter(c)) break;
            value = checked(value * 26 + char.ToUpperInvariant(c) - 'A' + 1);
        }
        return Math.Max(0, value - 1);
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

    private static FileStream OpenReadOnly(string path) => new(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);

    private static void ValidatePath(string filePath)
    {
        if (!File.Exists(filePath)) throw new FileNotFoundException("The selected workbook could not be found.", filePath);
        if (!string.Equals(Path.GetExtension(filePath), ".xlsx", StringComparison.OrdinalIgnoreCase))
            throw new NotSupportedException("Only .xlsx workbooks are supported.");
    }
}
