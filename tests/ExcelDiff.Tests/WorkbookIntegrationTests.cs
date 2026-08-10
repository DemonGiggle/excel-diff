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
    public void Reader_LoadsCompleteUsedRangeAndTypedValuesWithoutExcel()
    {
        var path = Path.Combine(_directory, "input.xlsx");
        using (var workbook = new XLWorkbook())
        {
            var sheet = workbook.Worksheets.Add("People");
            sheet.Cell("A1").Value = "ID";
            sheet.Cell("B1").Value = "Name";
            sheet.Cell("A2").Value = 42;
            sheet.Cell("B2").Value = "陳美玲";
            sheet.Cell("AD10").Value = new DateTime(2026, 8, 7);
            sheet.Cell("AD10").Style.DateFormat.Format = "yyyy-mm-dd";
            workbook.SaveAs(path);
        }

        var grid = new WorkbookReader().ReadGrid(path, "People", CancellationToken.None);

        Assert.Equal(10, grid.MaxRow);
        Assert.Equal(30, grid.MaxColumn);
        Assert.Equal("ID", grid.CellAt(1, 0).DisplayValue);
        Assert.Equal(CellValueKind.Number, grid.CellAt(2, 0).Kind);
        Assert.Equal("陳美玲", grid.CellAt(2, 1).DisplayValue);
        Assert.Equal(CellValueKind.Blank, grid.CellAt(5, 14).Kind);
        Assert.Equal(CellValueKind.Date, grid.CellAt(10, 29).Kind);
    }

    [Fact]
    public void Reader_LoadsLegacyXlsWithoutExcelInstalled()
    {
        var path = Path.Combine(AppContext.BaseDirectory, "TestData", "10x10.xls");
        var reader = new WorkbookReader();
        var sheets = reader.GetSheetNames(path);

        var grid = reader.ReadGrid(path, sheets[0], CancellationToken.None);

        Assert.NotEmpty(sheets);
        Assert.Equal(10, grid.MaxRow);
        Assert.Equal(10, grid.MaxColumn);
        Assert.NotEqual(CellValueKind.Blank, grid.CellAt(1, 0).Kind);
        Assert.Contains(grid.ReadIssues, issue => issue.Category == "Legacy workbook");
    }

    [Fact]
    public void Reader_PreservesLegacyXlsValueTypes()
    {
        var path = Path.Combine(AppContext.BaseDirectory, "TestData", "NumDoubleDateBoolString.xls");
        var reader = new WorkbookReader();
        var sheetName = reader.GetSheetNames(path)[0];

        var grid = reader.ReadGrid(path, sheetName, CancellationToken.None);
        var kinds = grid.Rows.Values.SelectMany(row => row.Cells.Values).Select(cell => cell.Kind).ToHashSet();

        Assert.Contains(CellValueKind.Text, kinds);
        Assert.Contains(CellValueKind.Number, kinds);
        Assert.Contains(CellValueKind.Date, kinds);
        Assert.Contains(CellValueKind.Boolean, kinds);
    }

    public void Dispose()
    {
        if (Directory.Exists(_directory)) Directory.Delete(_directory, true);
    }
}
