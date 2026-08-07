using ExcelDiff.Models;

namespace ExcelDiff.Services;

public interface IComparisonEngine
{
    ComparisonResult Compare(WorksheetData oldSheet, WorksheetData newSheet, ComparisonConfiguration configuration, CancellationToken cancellationToken, IProgress<int>? progress = null);
}

public sealed class ComparisonEngine : IComparisonEngine
{
    public ComparisonResult Compare(
        WorksheetData oldSheet,
        WorksheetData newSheet,
        ComparisonConfiguration configuration,
        CancellationToken cancellationToken,
        IProgress<int>? progress = null)
    {
        var mappings = configuration.Mappings.Where(m => m.IsIncluded && m.NewColumn is not null).ToArray();
        var keyMappings = mappings.Where(m => m.IsKey).ToArray();
        var issues = oldSheet.ReadIssues.Concat(newSheet.ReadIssues).ToList();
        if (mappings.Length == 0)
            throw Validation("Select at least one mapped field to compare.", "Mapping", "No fields are included.");
        if (keyMappings.Length == 0)
            throw Validation("Select at least one key field.", "Key", "No key field is selected.");
        if (mappings.Select(m => m.NewColumn!.Index).Distinct().Count() != mappings.Length)
            throw Validation("Each newer field can only be mapped once.", "Mapping", "Two or more older fields are mapped to the same newer field.");

        AddMappingIssues(oldSheet, newSheet, configuration, issues);
        AddKeyTypeWarnings(oldSheet, newSheet, keyMappings, issues);

        var oldIndex = BuildIndex(oldSheet, keyMappings, configuration.StrictTextComparison, "Old workbook", cancellationToken);
        progress?.Report(15);
        var newIndex = BuildIndex(newSheet, keyMappings, configuration.StrictTextComparison, "New workbook", cancellationToken);
        progress?.Report(30);

        var validationIssues = oldIndex.Issues.Concat(newIndex.Issues).ToArray();
        if (validationIssues.Length > 0)
            throw new ComparisonValidationException("Blank or duplicate keys must be fixed before comparison.", validationIssues);

        var results = new List<RowDifference>(oldIndex.Rows.Count + newIndex.Rows.Count);
        var allKeys = oldIndex.Rows.Keys.Union(newIndex.Rows.Keys, StringComparer.Ordinal).OrderBy(k => k, StringComparer.Ordinal).ToArray();
        for (var index = 0; index < allKeys.Length; index++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var stableKey = allKeys[index];
            var hasOld = oldIndex.Rows.TryGetValue(stableKey, out var oldEntry);
            var hasNew = newIndex.Rows.TryGetValue(stableKey, out var newEntry);
            var key = hasOld ? oldEntry!.Key : newEntry!.Key;
            var oldValues = hasOld ? ValuesFor(oldEntry!.Row, mappings, true) : EmptyValues(mappings);
            var newValues = hasNew ? ValuesFor(newEntry!.Row, mappings, false) : EmptyValues(mappings);

            if (!hasOld)
            {
                results.Add(Create(key, DifferenceStatus.Added, [], oldValues, newValues));
            }
            else if (!hasNew)
            {
                results.Add(Create(key, DifferenceStatus.Removed, [], oldValues, newValues));
            }
            else
            {
                var changes = new List<CellDifference>();
                foreach (var mapping in mappings.Where(m => !m.IsKey))
                {
                    var oldCell = oldEntry!.Row.CellAt(mapping.OldColumn.Index);
                    var newCell = newEntry!.Row.CellAt(mapping.NewColumn!.Index);
                    if (!Equivalent(oldCell, newCell, configuration.StrictTextComparison))
                        changes.Add(new CellDifference(mapping.OldHeader, oldCell.DisplayValue, newCell.DisplayValue));
                }
                results.Add(Create(key, changes.Count == 0 ? DifferenceStatus.Unchanged : DifferenceStatus.Changed, changes, oldValues, newValues));
            }

            if (index % 100 == 0) progress?.Report(30 + (allKeys.Length == 0 ? 70 : index * 70 / allKeys.Length));
        }

        var summary = new ComparisonSummary(
            results.Count(r => r.Status == DifferenceStatus.Unchanged),
            results.Count(r => r.Status == DifferenceStatus.Changed),
            results.Count(r => r.Status == DifferenceStatus.Added),
            results.Count(r => r.Status == DifferenceStatus.Removed),
            issues.Count(i => i.Severity != IssueSeverity.Information));
        progress?.Report(100);
        return new ComparisonResult(configuration, results, issues, summary, DateTimeOffset.Now);
    }

    private static void AddMappingIssues(WorksheetData oldSheet, WorksheetData newSheet, ComparisonConfiguration configuration, ICollection<ComparisonIssue> issues)
    {
        foreach (var mapping in configuration.Mappings.Where(m => m.NewColumn is null))
            issues.Add(new ComparisonIssue(IssueSeverity.Warning, "Mapping", $"Old field '{mapping.OldHeader}' is not mapped and will be ignored."));
        var mappedNew = configuration.Mappings.Where(m => m.NewColumn is not null).Select(m => m.NewColumn!.Index).ToHashSet();
        foreach (var header in newSheet.Headers.Where(h => !mappedNew.Contains(h.Index)))
            issues.Add(new ComparisonIssue(IssueSeverity.Warning, "Mapping", $"New field '{header.Name}' is not mapped and will be ignored."));
    }

    private static void AddKeyTypeWarnings(WorksheetData oldSheet, WorksheetData newSheet, IEnumerable<ColumnMapping> keyMappings, ICollection<ComparisonIssue> issues)
    {
        foreach (var mapping in keyMappings)
        {
            var oldKinds = oldSheet.Rows.Select(r => r.CellAt(mapping.OldColumn.Index).Kind).Where(k => k != CellValueKind.Blank).Distinct().ToArray();
            var newKinds = newSheet.Rows.Select(r => r.CellAt(mapping.NewColumn!.Index).Kind).Where(k => k != CellValueKind.Blank).Distinct().ToArray();
            if (oldKinds.Length > 1 || newKinds.Length > 1 || oldKinds.Length == 1 && newKinds.Length == 1 && oldKinds[0] != newKinds[0])
                issues.Add(new ComparisonIssue(IssueSeverity.Warning, "Key type", $"Key field '{mapping.OldHeader}' contains incompatible or mixed value types; text and numeric keys do not match."));
        }
    }

    private static (Dictionary<string, (CompositeRowKey Key, SheetRow Row)> Rows, List<ComparisonIssue> Issues) BuildIndex(
        WorksheetData sheet,
        IReadOnlyList<ColumnMapping> keys,
        bool strictText,
        string sourceName,
        CancellationToken cancellationToken)
    {
        var rows = new Dictionary<string, (CompositeRowKey, SheetRow)>(StringComparer.Ordinal);
        var issues = new List<ComparisonIssue>();
        foreach (var row in sheet.Rows)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var keyCells = keys.Select(k => row.CellAt(sourceName.StartsWith("Old", StringComparison.Ordinal) ? k.OldColumn.Index : k.NewColumn!.Index)).ToArray();
            if (keyCells.Any(c => c.Kind == CellValueKind.Blank))
            {
                issues.Add(new ComparisonIssue(IssueSeverity.Error, "Blank key", $"{sourceName} has a blank key value.", $"{sheet.SheetName}!Row {row.RowNumber}"));
                continue;
            }

            var displayKey = new CompositeRowKey(keyCells.Select(c => c.DisplayValue).ToArray());
            var stable = string.Join("\u001F", keyCells.Select(c => Comparable(c, strictText, true)));
            if (!rows.TryAdd(stable, (displayKey, row)))
                issues.Add(new ComparisonIssue(IssueSeverity.Error, "Duplicate key", $"{sourceName} contains duplicate key '{displayKey.Display}'.", $"{sheet.SheetName}!Row {row.RowNumber}"));
        }
        return (rows, issues);
    }

    private static bool Equivalent(CellData oldCell, CellData newCell, bool strictText) =>
        oldCell.Kind == newCell.Kind && string.Equals(Comparable(oldCell, strictText, false), Comparable(newCell, strictText, false), StringComparison.Ordinal);

    private static string Comparable(CellData cell, bool strictText, bool includeType) =>
        (includeType ? cell.Kind + ":" : string.Empty) + (cell.Kind == CellValueKind.Text && !strictText
            ? cell.CanonicalValue.Trim().ToUpperInvariant()
            : cell.CanonicalValue);

    private static Dictionary<string, string> ValuesFor(SheetRow row, IEnumerable<ColumnMapping> mappings, bool old) =>
        mappings.ToDictionary(m => m.OldHeader, m => row.CellAt(old ? m.OldColumn.Index : m.NewColumn!.Index).DisplayValue, StringComparer.OrdinalIgnoreCase);

    private static Dictionary<string, string> EmptyValues(IEnumerable<ColumnMapping> mappings) =>
        mappings.ToDictionary(m => m.OldHeader, _ => string.Empty, StringComparer.OrdinalIgnoreCase);

    private static RowDifference Create(CompositeRowKey key, DifferenceStatus status, IReadOnlyList<CellDifference> changes,
        IReadOnlyDictionary<string, string> oldValues, IReadOnlyDictionary<string, string> newValues) =>
        new() { Key = key, Status = status, Changes = changes, OldValues = oldValues, NewValues = newValues };

    private static ComparisonValidationException Validation(string message, string category, string detail) =>
        new(message, [new ComparisonIssue(IssueSeverity.Error, category, detail)]);
}
