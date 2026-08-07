using ClosedXML.Excel;
using ExcelDiff.Models;
using System.IO;

namespace ExcelDiff.Services;

public interface IExcelReportExporter
{
    void Export(ComparisonResult result, string outputPath, CancellationToken cancellationToken);
}

public sealed class ExcelReportExporter : IExcelReportExporter
{
    private static readonly XLColor Navy = XLColor.FromHtml("#172033");
    private static readonly XLColor Blue = XLColor.FromHtml("#315EFB");
    private static readonly XLColor PaleBlue = XLColor.FromHtml("#EAF0FF");
    private static readonly XLColor PaleGreen = XLColor.FromHtml("#E8F7EE");
    private static readonly XLColor PaleRed = XLColor.FromHtml("#FDECEC");
    private static readonly XLColor PaleAmber = XLColor.FromHtml("#FFF4D8");

    public void Export(ComparisonResult result, string outputPath, CancellationToken cancellationToken)
    {
        ArgumentNullException.ThrowIfNull(result);
        if (!string.Equals(Path.GetExtension(outputPath), ".xlsx", StringComparison.OrdinalIgnoreCase))
            throw new NotSupportedException("The report must be saved as an .xlsx file.");

        using var workbook = new XLWorkbook();
        WriteSummary(workbook, result);
        cancellationToken.ThrowIfCancellationRequested();
        WriteChangedCells(workbook, result);
        cancellationToken.ThrowIfCancellationRequested();
        WriteRows(workbook, "Added Rows", result, DifferenceStatus.Added, PaleGreen);
        WriteRows(workbook, "Removed Rows", result, DifferenceStatus.Removed, PaleRed);
        WriteIssues(workbook, result);
        cancellationToken.ThrowIfCancellationRequested();
        workbook.SaveAs(outputPath, validate: true);
    }

    private static void WriteSummary(XLWorkbook workbook, ComparisonResult result)
    {
        var ws = workbook.Worksheets.Add("Summary");
        ws.Cell("A1").Value = "Excel comparison report";
        ws.Range("A1:D1").Merge();
        ws.Range("A1:D1").Style.Fill.BackgroundColor = Navy;
        ws.Range("A1:D1").Style.Font.FontColor = XLColor.White;
        ws.Range("A1:D1").Style.Font.Bold = true;
        ws.Range("A1:D1").Style.Font.FontSize = 18;
        ws.Row(1).Height = 32;

        var details = new (string Label, string Value)[]
        {
            ("Old workbook", Path.GetFileName(result.Configuration.OldFilePath)),
            ("Old sheet", result.Configuration.OldSheetName),
            ("New workbook", Path.GetFileName(result.Configuration.NewFilePath)),
            ("New sheet", result.Configuration.NewSheetName),
            ("Compared at", result.CompletedAt.LocalDateTime.ToString("yyyy-MM-dd HH:mm:ss")),
            ("Text comparison", result.Configuration.StrictTextComparison ? "Strict" : "Ignore leading/trailing spaces and letter case"),
            ("Key fields", string.Join(", ", result.Configuration.Mappings.Where(m => m.IsKey).Select(m => m.OldHeader))),
            ("Included fields", result.Configuration.Mappings.Count(m => m.IsIncluded && m.NewColumn is not null).ToString()),
        };
        var row = 3;
        foreach (var detail in details)
        {
            ws.Cell(row, 1).Value = detail.Label;
            ws.Cell(row, 2).Value = detail.Value;
            ws.Cell(row, 1).Style.Font.Bold = true;
            row++;
        }

        row += 1;
        ws.Cell(row, 1).Value = "Status";
        ws.Cell(row, 2).Value = "Rows";
        StyleHeader(ws.Range(row, 1, row, 2));
        var statuses = new (string Name, int Count, XLColor Color)[]
        {
            ("Unchanged", result.Summary.Unchanged, PaleBlue),
            ("Changed", result.Summary.Changed, PaleAmber),
            ("Added", result.Summary.Added, PaleGreen),
            ("Removed", result.Summary.Removed, PaleRed),
            ("Warnings / problems", result.Summary.Problems, PaleAmber)
        };
        foreach (var status in statuses)
        {
            row++;
            ws.Cell(row, 1).Value = status.Name;
            ws.Cell(row, 2).Value = status.Count;
            ws.Range(row, 1, row, 2).Style.Fill.BackgroundColor = status.Color;
        }
        ws.Column(1).Width = 24;
        ws.Column(2).Width = 72;
        ws.Column(2).Style.Alignment.WrapText = true;
        ws.SheetView.FreezeRows(1);
    }

    private static void WriteChangedCells(XLWorkbook workbook, ComparisonResult result)
    {
        var ws = workbook.Worksheets.Add("Changed Cells");
        var keyHeaders = result.Configuration.Mappings.Where(m => m.IsKey).Select(m => m.OldHeader).ToArray();
        var headers = keyHeaders.Concat(["Field", "Old value", "New value"]).ToArray();
        WriteHeaderRow(ws, headers);
        var row = 2;
        foreach (var difference in result.Rows.Where(r => r.Status == DifferenceStatus.Changed))
        {
            foreach (var change in difference.Changes)
            {
                for (var key = 0; key < keyHeaders.Length; key++) ws.Cell(row, key + 1).Value = difference.Key.Parts.ElementAtOrDefault(key) ?? string.Empty;
                ws.Cell(row, keyHeaders.Length + 1).Value = change.FieldName;
                ws.Cell(row, keyHeaders.Length + 2).Value = change.OldValue;
                ws.Cell(row, keyHeaders.Length + 3).Value = change.NewValue;
                ws.Range(row, 1, row, headers.Length).Style.Fill.BackgroundColor = PaleAmber;
                row++;
            }
        }
        FinalizeTable(ws, headers.Length, row - 1);
    }

    private static void WriteRows(XLWorkbook workbook, string sheetName, ComparisonResult result, DifferenceStatus status, XLColor fill)
    {
        var ws = workbook.Worksheets.Add(sheetName);
        var mappings = result.Configuration.Mappings.Where(m => m.IsIncluded && m.NewColumn is not null).ToArray();
        WriteHeaderRow(ws, mappings.Select(m => m.OldHeader).ToArray());
        var row = 2;
        foreach (var difference in result.Rows.Where(r => r.Status == status))
        {
            var values = status == DifferenceStatus.Added ? difference.NewValues : difference.OldValues;
            for (var column = 0; column < mappings.Length; column++)
                ws.Cell(row, column + 1).Value = values.GetValueOrDefault(mappings[column].OldHeader, string.Empty);
            ws.Range(row, 1, row, mappings.Length).Style.Fill.BackgroundColor = fill;
            row++;
        }
        FinalizeTable(ws, Math.Max(1, mappings.Length), row - 1);
    }

    private static void WriteIssues(XLWorkbook workbook, ComparisonResult result)
    {
        var ws = workbook.Worksheets.Add("Issues");
        WriteHeaderRow(ws, ["Severity", "Category", "Message", "Location"]);
        var row = 2;
        foreach (var issue in result.Issues)
        {
            ws.Cell(row, 1).Value = issue.Severity.ToString();
            ws.Cell(row, 2).Value = issue.Category;
            ws.Cell(row, 3).Value = issue.Message;
            ws.Cell(row, 4).Value = issue.Location ?? string.Empty;
            ws.Range(row, 1, row, 4).Style.Fill.BackgroundColor = issue.Severity == IssueSeverity.Error ? PaleRed : PaleAmber;
            row++;
        }
        FinalizeTable(ws, 4, row - 1);
    }

    private static void WriteHeaderRow(IXLWorksheet ws, IReadOnlyList<string> headers)
    {
        for (var index = 0; index < headers.Count; index++) ws.Cell(1, index + 1).Value = headers[index];
        StyleHeader(ws.Range(1, 1, 1, Math.Max(1, headers.Count)));
        ws.SheetView.FreezeRows(1);
    }

    private static void StyleHeader(IXLRange range)
    {
        range.Style.Fill.BackgroundColor = Blue;
        range.Style.Font.FontColor = XLColor.White;
        range.Style.Font.Bold = true;
        range.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
    }

    private static void FinalizeTable(IXLWorksheet ws, int lastColumn, int lastDataRow)
    {
        if (lastDataRow >= 2) ws.Range(1, 1, lastDataRow, lastColumn).SetAutoFilter();
        ws.Columns(1, lastColumn).AdjustToContents(1, Math.Max(1, Math.Min(lastDataRow, 500)));
        foreach (var column in ws.Columns(1, lastColumn))
        {
            if (column.Width > 48) column.Width = 48;
            if (column.Width < 10) column.Width = 10;
            column.Style.Alignment.WrapText = true;
            column.Style.Alignment.Vertical = XLAlignmentVerticalValues.Top;
        }
    }
}
