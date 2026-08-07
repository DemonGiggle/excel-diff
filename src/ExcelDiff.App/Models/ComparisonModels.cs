using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace ExcelDiff.Models;

public enum CellValueKind { Blank, Text, Number, Date, Boolean, Error }
public enum DifferenceStatus { Unchanged, Changed, Added, Removed }
public enum IssueSeverity { Information, Warning, Error }

public sealed record CellData(
    CellValueKind Kind,
    string DisplayValue,
    string CanonicalValue,
    bool HasFormula = false,
    bool MissingCachedFormulaValue = false)
{
    public static readonly CellData Blank = new(CellValueKind.Blank, string.Empty, string.Empty);
}

public sealed record SheetHeader(int Index, string Name);

public sealed record SheetRow(int RowNumber, IReadOnlyList<CellData> Cells)
{
    public CellData CellAt(int index) => index >= 0 && index < Cells.Count ? Cells[index] : CellData.Blank;
}

public sealed record WorksheetData(
    string FilePath,
    string SheetName,
    int HeaderRow,
    IReadOnlyList<SheetHeader> Headers,
    IReadOnlyList<SheetRow> Rows,
    IReadOnlyList<ComparisonIssue> ReadIssues);

public sealed class ColumnMapping : INotifyPropertyChanged
{
    private SheetHeader? _newColumn;
    private bool _isIncluded = true;
    private bool _isKey;

    public required SheetHeader OldColumn { get; init; }
    public required IReadOnlyList<SheetHeader> NewColumnOptions { get; init; }
    public SheetHeader? NewColumn { get => _newColumn; set { _newColumn = value; OnPropertyChanged(); } }
    public bool IsIncluded { get => _isIncluded; set { _isIncluded = value; if (!value) IsKey = false; OnPropertyChanged(); } }
    public bool IsKey { get => _isKey; set { _isKey = value; if (value) IsIncluded = true; OnPropertyChanged(); } }
    public string OldHeader => OldColumn.Name;

    public event PropertyChangedEventHandler? PropertyChanged;
    private void OnPropertyChanged([CallerMemberName] string? name = null) =>
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
}

public sealed record ComparisonConfiguration(
    string OldFilePath,
    string NewFilePath,
    string OldSheetName,
    string NewSheetName,
    int OldHeaderRow,
    int NewHeaderRow,
    IReadOnlyList<ColumnMapping> Mappings,
    bool StrictTextComparison);

public sealed record CompositeRowKey(IReadOnlyList<string> Parts)
{
    public string Display => string.Join(" | ", Parts);
    public string StableValue => string.Join("\u001F", Parts);
}

public sealed record CellDifference(string FieldName, string OldValue, string NewValue)
{
    public string ChangeText => $"{OldValue}  →  {NewValue}";
}

public sealed record ComparisonIssue(IssueSeverity Severity, string Category, string Message, string? Location = null);

public sealed class RowDifference
{
    public required CompositeRowKey Key { get; init; }
    public required DifferenceStatus Status { get; init; }
    public required IReadOnlyList<CellDifference> Changes { get; init; }
    public required IReadOnlyDictionary<string, string> OldValues { get; init; }
    public required IReadOnlyDictionary<string, string> NewValues { get; init; }
    public string KeyDisplay => Key.Display;
    public string StatusText => Status.ToString();
    public string ChangeSummary => Status switch
    {
        DifferenceStatus.Changed => $"{Changes.Count} field{(Changes.Count == 1 ? "" : "s")} changed",
        DifferenceStatus.Added => "Row added",
        DifferenceStatus.Removed => "Row removed",
        _ => "No changes"
    };
    public string SearchText => string.Join(" ", new[] { Key.Display, StatusText }
        .Concat(Changes.SelectMany(c => new[] { c.FieldName, c.OldValue, c.NewValue }))
        .Concat(OldValues.Values).Concat(NewValues.Values));
}

public sealed record ComparisonSummary(int Unchanged, int Changed, int Added, int Removed, int Problems)
{
    public int Total => Unchanged + Changed + Added + Removed;
}

public sealed record ComparisonResult(
    ComparisonConfiguration Configuration,
    IReadOnlyList<RowDifference> Rows,
    IReadOnlyList<ComparisonIssue> Issues,
    ComparisonSummary Summary,
    DateTimeOffset CompletedAt);

public sealed class ComparisonValidationException : Exception
{
    public ComparisonValidationException(string message, IReadOnlyList<ComparisonIssue> issues) : base(message) => Issues = issues;
    public IReadOnlyList<ComparisonIssue> Issues { get; }
}
