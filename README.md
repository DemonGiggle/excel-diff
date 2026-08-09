# Excel Compare

Excel Compare is an offline Windows desktop tool for comparing one table from two `.xlsx` or legacy `.xls` workbooks. It matches rows by business keys instead of row position, presents an accessible interactive review, and exports a shareable `.xlsx` report.

![Excel Compare showing an employee roster comparison with changed, added, removed, and unchanged rows](docs/images/excel-compare-demo.png)

## Try the demo

Download the [older employee roster](outputs/readme-demo/employee-roster-old.xlsx) and [newer employee roster](outputs/readme-demo/employee-roster-new.xlsx), then select `Employee ID` as the row key. The comparison contains five changed rows, two additions, one removal, and four unchanged rows.

## Use

1. Run `ExcelDiff.exe`.
2. Choose the older and newer `.xlsx` or `.xls` files.
3. Select a worksheet and header row from each file.
4. Review the automatic field mapping and mark one or more unique **Row key** fields.
5. Compare, filter the result, inspect individual rows, and export the report.

Source workbooks are opened read-only. The app never runs formulas, macros, or external links. Formula cells are compared using their last saved results.

## Build

Requirements: Windows 10/11 x64 and the .NET 8 SDK.

```powershell
.\build.ps1
```

The script locates either the repository-local SDK or an installed .NET 8 SDK, restores packages, runs the tests, and builds the portable application. The application folder and ZIP are written to `artifacts`.

To build without running tests:

```powershell
.\build.ps1 -SkipTests
```

## Supported scope

- `.xlsx` and legacy `.xls` input files
- One selected sheet/table pair per comparison
- Single or composite row keys
- Text, number, date, Boolean, blank, error, and saved formula-result values
- Formatting, charts, comments, macros, and workbook structure are intentionally ignored
