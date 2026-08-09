using ClosedXML.Excel;
using ExcelDiff.Models;
using ExcelDiff.Services;
using System.IO;
using Xunit;

namespace ExcelDiff.Tests;

public sealed class WorkbookIntegrationTests : IDisposable
{
    private readonly string _directory = Path.Combine(Path.GetTempPath(), $"ExcelDiffTests-{Guid.NewGuid():N}");

    public WorkbookIntegrationTests() => Directory.CreateDirectory(_directory);

    [Fact]
    public void Reader_LoadsTypedValuesAndDoesNotNeedExcel()
    {
        var path = Path.Combine(_directory, "input.xlsx");
        using (var workbook = new XLWorkbook())
        {
            var ws = workbook.Worksheets.Add("People");
            ws.Cell("A1").Value = "ID";
            ws.Cell("B1").Value = "Name";
            ws.Cell("C1").Value = "Started";
            ws.Cell("A2").Value = 42;
            ws.Cell("B2").Value = "林怡君";
            ws.Cell("C2").Value = new DateTime(2026, 8, 7);
            ws.Cell("C2").Style.DateFormat.Format = "yyyy-mm-dd";
            workbook.SaveAs(path);
        }

        var reader = new WorkbookReader();
        var sheet = reader.ReadSheet(path, "People", 1, CancellationToken.None);

        Assert.Equal(["ID", "Name", "Started"], sheet.Headers.Select(h => h.Name));
        Assert.Equal(CellValueKind.Number, sheet.Rows[0].CellAt(0).Kind);
        Assert.Equal("林怡君", sheet.Rows[0].CellAt(1).DisplayValue);
        Assert.Equal(CellValueKind.Date, sheet.Rows[0].CellAt(2).Kind);
    }

    [Fact]
    public void Reader_LoadsLegacyXlsWithoutExcelInstalled()
    {
        var path = Path.Combine(AppContext.BaseDirectory, "TestData", "10x10.xls");
        var reader = new WorkbookReader();

        var sheets = reader.GetSheetNames(path);
        var sheet = reader.ReadSheet(path, sheets[0], 1, CancellationToken.None);

        Assert.NotEmpty(sheets);
        Assert.Equal(9, sheet.Headers.Count);
        Assert.Equal(9, sheet.Rows.Count);
        Assert.NotEqual(CellValueKind.Blank, sheet.Rows[0].CellAt(0).Kind);
        Assert.Contains(sheet.ReadIssues, issue => issue.Category == "Legacy workbook");
    }

    [Fact]
    public void Reader_PreservesLegacyXlsValueTypes()
    {
        var path = Path.Combine(AppContext.BaseDirectory, "TestData", "NumDoubleDateBoolString.xls");
        var reader = new WorkbookReader();
        var sheetName = reader.GetSheetNames(path)[0];

        var sheet = reader.ReadSheet(path, sheetName, 1, CancellationToken.None);
        var kinds = sheet.Rows.SelectMany(row => row.Cells).Select(cell => cell.Kind).ToHashSet();

        Assert.Contains(CellValueKind.Text, kinds);
        Assert.Contains(CellValueKind.Number, kinds);
        Assert.Contains(CellValueKind.Date, kinds);
        Assert.Contains(CellValueKind.Boolean, kinds);
    }

    [Fact]
    public void Exporter_CreatesAllExpectedReportSheets()
    {
        var mapping = new[]
        {
            new ColumnMapping { OldColumn = new SheetHeader(0, "ID"), NewColumnOptions = [new SheetHeader(0, "ID")], NewColumn = new SheetHeader(0, "ID"), IsKey = true },
            new ColumnMapping { OldColumn = new SheetHeader(1, "Name"), NewColumnOptions = [new SheetHeader(1, "Name")], NewColumn = new SheetHeader(1, "Name") }
        };
        var config = new ComparisonConfiguration("old.xlsx", "new.xlsx", "Data", "Data", 1, 1, mapping, false);
        var row = new RowDifference
        {
            Key = new CompositeRowKey(["A1"]), Status = DifferenceStatus.Changed,
            Changes = [new CellDifference("Name", "Alice", "Alicia")],
            OldValues = new Dictionary<string, string> { ["ID"] = "A1", ["Name"] = "Alice" },
            NewValues = new Dictionary<string, string> { ["ID"] = "A1", ["Name"] = "Alicia" }
        };
        var result = new ComparisonResult(config, [row], [new ComparisonIssue(IssueSeverity.Warning, "Test", "Example")], new ComparisonSummary(0, 1, 0, 0, 1), DateTimeOffset.Now);
        var path = Path.Combine(_directory, "report.xlsx");

        new ExcelReportExporter().Export(result, path, CancellationToken.None);

        using var workbook = new XLWorkbook(path);
        Assert.Equal(["Summary", "Changed Cells", "Added Rows", "Removed Rows", "Issues"], workbook.Worksheets.Select(w => w.Name));
        Assert.Equal("Alice", workbook.Worksheet("Changed Cells").Cell("C2").GetString());
        Assert.Equal("Alicia", workbook.Worksheet("Changed Cells").Cell("D2").GetString());
    }

    public void Dispose()
    {
        if (Directory.Exists(_directory)) Directory.Delete(_directory, true);
    }
}
