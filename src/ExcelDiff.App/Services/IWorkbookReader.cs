using ExcelDiff.Models;

namespace ExcelDiff.Services;

public interface IWorkbookReader
{
    IReadOnlyList<string> GetSheetNames(string filePath);
    WorksheetData ReadSheet(string filePath, string sheetName, int headerRow, CancellationToken cancellationToken, IProgress<int>? progress = null);
}
