using ExcelDiff.Models;
using System.IO;

namespace ExcelDiff.Services;

public sealed class WorkbookReader : IWorkbookReader
{
    private readonly IWorkbookReader _openXmlReader = new OpenXmlWorkbookReader();
    private readonly IWorkbookReader _binaryReader = new BinaryExcelWorkbookReader();

    public IReadOnlyList<string> GetSheetNames(string filePath) => ReaderFor(filePath).GetSheetNames(filePath);

    public WorksheetGrid ReadGrid(
        string filePath,
        string sheetName,
        CancellationToken cancellationToken,
        IProgress<int>? progress = null) =>
        ReaderFor(filePath).ReadGrid(filePath, sheetName, cancellationToken, progress);

    private IWorkbookReader ReaderFor(string filePath)
    {
        var extension = Path.GetExtension(filePath);
        if (string.Equals(extension, ".xlsx", StringComparison.OrdinalIgnoreCase)) return _openXmlReader;
        if (string.Equals(extension, ".xls", StringComparison.OrdinalIgnoreCase)) return _binaryReader;
        throw new NotSupportedException("Only .xlsx and .xls workbooks are supported.");
    }
}
