using ExcelDiff.Models;
using ExcelDiff.Services;
using Xunit;

namespace ExcelDiff.Tests;

public sealed class ComparisonEngineTests
{
    private readonly ComparisonEngine _engine = new();

    [Fact]
    public void ReorderedRows_WithSameKeys_AreUnchanged()
    {
        var oldSheet = Sheet(Row(2, "A1", "Alice"), Row(3, "B2", "Bob"));
        var newSheet = Sheet(Row(2, "B2", "Bob"), Row(3, "A1", "Alice"));

        var result = Compare(oldSheet, newSheet);

        Assert.Equal(2, result.Summary.Unchanged);
        Assert.Equal(0, result.Summary.Changed);
    }

    [Fact]
    public void ChangedAddedAndRemovedRows_AreClassified()
    {
        var oldSheet = Sheet(Row(2, "A1", "Alice"), Row(3, "B2", "Bob"));
        var newSheet = Sheet(Row(2, "A1", "Alicia"), Row(3, "C3", "Chen"));

        var result = Compare(oldSheet, newSheet);

        Assert.Equal(1, result.Summary.Changed);
        Assert.Equal(1, result.Summary.Added);
        Assert.Equal(1, result.Summary.Removed);
        Assert.Equal("Name", result.Rows.Single(r => r.Status == DifferenceStatus.Changed).Changes.Single().FieldName);
    }

    [Fact]
    public void DefaultTextComparison_IgnoresCaseAndSurroundingSpaces()
    {
        var result = Compare(Sheet(Row(2, "A1", " Alice ")), Sheet(Row(2, "a1", "alice")));
        Assert.Equal(1, result.Summary.Unchanged);
    }

    [Fact]
    public void StrictTextComparison_FlagsCaseAndSpaces()
    {
        var oldSheet = Sheet(Row(2, "A1", " Alice "));
        var newSheet = Sheet(Row(2, "A1", "alice"));
        var result = Compare(oldSheet, newSheet, strict: true);
        Assert.Equal(1, result.Summary.Changed);
    }

    [Fact]
    public void DuplicateKeys_BlockComparison()
    {
        var oldSheet = Sheet(Row(2, "A1", "Alice"), Row(3, "A1", "Another"));
        var exception = Assert.Throws<ComparisonValidationException>(() => Compare(oldSheet, Sheet(Row(2, "A1", "Alice"))));
        Assert.Contains(exception.Issues, issue => issue.Category == "Duplicate key");
    }

    [Fact]
    public void BlankKeys_BlockComparison()
    {
        var oldSheet = Sheet(Row(2, null, "Alice"));
        var exception = Assert.Throws<ComparisonValidationException>(() => Compare(oldSheet, Sheet(Row(2, "A1", "Alice"))));
        Assert.Contains(exception.Issues, issue => issue.Category == "Blank key");
    }

    [Fact]
    public void NumericAndTextKeys_DoNotMatch_AndProduceWarning()
    {
        var oldSheet = Sheet(new SheetRow(2, [Number(1), Text("Alice")]));
        var newSheet = Sheet(Row(2, "1", "Alice"));
        var result = Compare(oldSheet, newSheet);
        Assert.Equal(1, result.Summary.Added);
        Assert.Equal(1, result.Summary.Removed);
        Assert.Contains(result.Issues, issue => issue.Category == "Key type");
    }

    [Fact]
    public void CompositeKeys_MatchRowsIndependentlyOfOrder()
    {
        var headers = new[] { new SheetHeader(0, "Company"), new SheetHeader(1, "Employee ID"), new SheetHeader(2, "Name") };
        var oldSheet = new WorksheetData("old.xlsx", "Data", 1, headers,
            [new SheetRow(2, [Text("North"), Text("7"), Text("Alice")]), new SheetRow(3, [Text("South"), Text("7"), Text("Bob")])], []);
        var newSheet = new WorksheetData("new.xlsx", "Data", 1, headers,
            [new SheetRow(2, [Text("South"), Text("7"), Text("Bob")]), new SheetRow(3, [Text("North"), Text("7"), Text("Alice")])], []);
        var mappings = headers.Select((header, index) => new ColumnMapping
        {
            OldColumn = header, NewColumnOptions = headers, NewColumn = headers[index], IsKey = index < 2
        }).ToArray();
        var config = new ComparisonConfiguration("old.xlsx", "new.xlsx", "Data", "Data", 1, 1, mappings, false);

        var result = _engine.Compare(oldSheet, newSheet, config, CancellationToken.None);

        Assert.Equal(2, result.Summary.Unchanged);
    }

    [Fact]
    public void DuplicateTargetMappings_BlockComparison()
    {
        var oldSheet = Sheet(Row(2, "A1", "Alice"));
        var newSheet = Sheet(Row(2, "A1", "Alice"));
        var mappings = new[]
        {
            new ColumnMapping { OldColumn = oldSheet.Headers[0], NewColumnOptions = newSheet.Headers, NewColumn = newSheet.Headers[0], IsKey = true },
            new ColumnMapping { OldColumn = oldSheet.Headers[1], NewColumnOptions = newSheet.Headers, NewColumn = newSheet.Headers[0] }
        };
        var config = new ComparisonConfiguration("old.xlsx", "new.xlsx", "Data", "Data", 1, 1, mappings, false);

        var exception = Assert.Throws<ComparisonValidationException>(() => _engine.Compare(oldSheet, newSheet, config, CancellationToken.None));

        Assert.Contains(exception.Issues, issue => issue.Category == "Mapping");
    }

    private ComparisonResult Compare(WorksheetData oldSheet, WorksheetData newSheet, bool strict = false)
    {
        var mappings = new[]
        {
            new ColumnMapping { OldColumn = oldSheet.Headers[0], NewColumnOptions = newSheet.Headers, NewColumn = newSheet.Headers[0], IsKey = true },
            new ColumnMapping { OldColumn = oldSheet.Headers[1], NewColumnOptions = newSheet.Headers, NewColumn = newSheet.Headers[1] }
        };
        var config = new ComparisonConfiguration("old.xlsx", "new.xlsx", "Data", "Data", 1, 1, mappings, strict);
        return _engine.Compare(oldSheet, newSheet, config, CancellationToken.None);
    }

    private static WorksheetData Sheet(params SheetRow[] rows) =>
        new("test.xlsx", "Data", 1, [new SheetHeader(0, "ID"), new SheetHeader(1, "Name")], rows, []);

    private static SheetRow Row(int number, string? id, string name) => new(number, [id is null ? CellData.Blank : Text(id), Text(name)]);
    private static CellData Text(string value) => new(CellValueKind.Text, value, value);
    private static CellData Number(decimal value) => new(CellValueKind.Number, value.ToString(), value.ToString());
}
