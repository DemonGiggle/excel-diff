using ExcelDiff.Models;

namespace ExcelDiff.Services;

public interface IWorkbookReader
{
    IReadOnlyList<string> GetSheetNames(string filePath);
    WorksheetGrid ReadGrid(string filePath, string sheetName, CancellationToken cancellationToken, IProgress<int>? progress = null);
}
