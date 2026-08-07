# Excel Compare

Excel Compare is an offline Windows desktop tool for comparing one table from two `.xlsx` workbooks. It matches rows by business keys instead of row position, presents an accessible interactive review, and exports a shareable Excel report.

## Use

1. Run `ExcelDiff.exe`.
2. Choose the older and newer `.xlsx` files.
3. Select a worksheet and header row from each file.
4. Review the automatic field mapping and mark one or more unique **Row key** fields.
5. Compare, filter the result, inspect individual rows, and export the report.

Source workbooks are opened read-only. The app never runs formulas, macros, or external links. Formula cells are compared using their last saved results.

## Build

Requirements: Windows 10/11 x64 and the .NET 8 SDK.

```powershell
dotnet restore ExcelDiff.sln
dotnet test ExcelDiff.sln -c Release
.\scripts\publish.ps1
```

The portable application and ZIP are written to `artifacts`.

## Supported scope

- `.xlsx` files only
- One selected sheet/table pair per comparison
- Single or composite row keys
- Text, number, date, Boolean, blank, error, and saved formula-result values
- Formatting, charts, comments, macros, and workbook structure are intentionally ignored
