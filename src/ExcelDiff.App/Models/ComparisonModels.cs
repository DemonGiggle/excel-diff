namespace ExcelDiff.Models;

public enum CellValueKind { Blank, Text, Number, Date, Boolean, Error }
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

public sealed record ComparisonIssue(IssueSeverity Severity, string Category, string Message, string? Location = null);
