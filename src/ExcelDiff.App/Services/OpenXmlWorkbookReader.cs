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

    public WorksheetData ReadSheet(
        string filePath,
        string sheetName,
        int headerRow,
        CancellationToken cancellationToken,
        IProgress<int>? progress = null)
    {
        ValidatePath(filePath);
        if (headerRow < 1) throw new ArgumentOutOfRangeException(nameof(headerRow), "Header row must be 1 or greater.");

        using var stream = OpenReadOnly(filePath);
        using var document = SpreadsheetDocument.Open(stream, false);
        var workbookPart = document.WorkbookPart ?? throw new InvalidDataException("The workbook structure is missing.");
        var sheets = workbookPart.Workbook?.Sheets ?? throw new InvalidDataException("The workbook has no worksheet list.");
        var sheet = sheets.Elements<Sheet>()
            .FirstOrDefault(s => string.Equals(s.Name?.Value, sheetName, StringComparison.Ordinal));
        if (sheet?.Id?.Value is null) throw new InvalidDataException($"Sheet '{sheetName}' was not found.");

        var worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id.Value);
        var sharedStrings = workbookPart.SharedStringTablePart?.SharedStringTable;
        var styles = workbookPart.WorkbookStylesPart?.Stylesheet;
        var issues = new List<ComparisonIssue>();
        var dataRows = new List<SheetRow>();
        List<SheetHeader>? headers = null;
        var estimatedLastRow = ReadEstimatedLastRow(worksheetPart);

        using var reader = OpenXmlReader.Create(worksheetPart);
        while (reader.Read())
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (!reader.IsStartElement || reader.ElementType != typeof(Row)) continue;
            if (reader.LoadCurrentElement() is not Row row) continue;
            var rowNumber = checked((int)(row.RowIndex?.Value ?? 0));
            if (rowNumber < headerRow) continue;

            var indexedCells = row.Elements<Cell>().ToDictionary(c => GetColumnIndex(c.CellReference?.Value));
            if (rowNumber == headerRow)
            {
                if (indexedCells.Count == 0)
                    throw new InvalidDataException($"Row {headerRow} in '{sheetName}' is empty. Choose the row containing column names.");

                var lastColumn = indexedCells.Keys.Max();
                headers = BuildHeaders(indexedCells, lastColumn, sharedStrings, styles, issues, sheetName);
                continue;
            }

            if (headers is null) continue;
            var values = new CellData[headers.Count];
            var hasAnyValue = false;
            for (var column = 0; column < headers.Count; column++)
            {
                var value = indexedCells.TryGetValue(column, out var cell)
                    ? ReadCell(cell, sharedStrings, styles)
                    : CellData.Blank;
                values[column] = value;
                hasAnyValue |= value.Kind != CellValueKind.Blank;
                if (value.MissingCachedFormulaValue)
                {
                    issues.Add(new ComparisonIssue(IssueSeverity.Warning, "Formula",
                        "A formula has no saved result. It is treated as blank because formulas are never executed.",
                        $"{sheetName}!{GetColumnName(column)}{rowNumber}"));
                }
            }

            if (hasAnyValue) dataRows.Add(new SheetRow(rowNumber, values));
            if (estimatedLastRow > 0 && rowNumber % 250 == 0)
                progress?.Report(Math.Clamp(rowNumber * 100 / estimatedLastRow, 0, 99));
        }

        if (headers is null)
            throw new InvalidDataException($"Header row {headerRow} was not found in '{sheetName}'.");

        if (dataRows.Count == 0)
            issues.Add(new ComparisonIssue(IssueSeverity.Warning, "Empty sheet", "No data rows were found below the selected header row.", sheetName));
        progress?.Report(100);
        return new WorksheetData(filePath, sheetName, headerRow, headers, dataRows, issues);
    }

    private static List<SheetHeader> BuildHeaders(
        IReadOnlyDictionary<int, Cell> cells,
        int lastColumn,
        SharedStringTable? sharedStrings,
        Stylesheet? styles,
        ICollection<ComparisonIssue> issues,
        string sheetName)
    {
        var result = new List<SheetHeader>(lastColumn + 1);
        var seen = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        for (var index = 0; index <= lastColumn; index++)
        {
            var text = cells.TryGetValue(index, out var cell) ? ReadCell(cell, sharedStrings, styles).DisplayValue.Trim() : string.Empty;
            if (string.IsNullOrWhiteSpace(text))
            {
                text = $"Unnamed column {GetColumnName(index)}";
                issues.Add(new ComparisonIssue(IssueSeverity.Warning, "Header", "A blank header was assigned a temporary name.", $"{sheetName}!{GetColumnName(index)}"));
            }

            var normalized = text.Trim();
            if (seen.TryGetValue(normalized, out _))
            {
                issues.Add(new ComparisonIssue(IssueSeverity.Warning, "Header", $"Duplicate header '{text}' was disambiguated.", $"{sheetName}!{GetColumnName(index)}"));
                text = $"{text} [{GetColumnName(index)}]";
            }
            else seen[normalized] = index;
            result.Add(new SheetHeader(index, text));
        }
        return result;
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
