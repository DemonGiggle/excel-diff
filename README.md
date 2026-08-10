# Excel Compare

Excel Compare is an offline Windows desktop tool that presents two `.xlsx` or legacy `.xls` worksheets as one clear visual diff. It detects inserted and removed rows or columns so one structural change does not create a cascade of false differences.

![Excel Compare unified worksheet view with removed values struck through in red and added values highlighted in green](docs/images/excel-compare-demo.png)

## Try the demo

Download the [older employee roster](outputs/readme-demo/employee-roster-old.xlsx) and [newer employee roster](outputs/readme-demo/employee-roster-new.xlsx), select the worksheet in each file, and compare them.

## Use

1. Run `ExcelDiff.exe`.
2. Choose the older and newer `.xlsx` or `.xls` files.
3. Select one worksheet from each file and choose **Compare worksheets**.
4. Review the unified worksheet: red struck-through values were removed, green values were added, and changed cells show both old and new values.
5. Use **Previous change** and **Next change** to move between changed rows. Consecutive unchanged rows are folded automatically.

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
- One selected worksheet pair per comparison
- Automatic used range from `A1` through the furthest populated cell
- Automatic inserted/removed row and column alignment
- Unified cell view with folded unchanged rows and changed-row navigation
- Text, number, date, Boolean, blank, error, and saved formula-result values
- Formatting, charts, comments, macros, and workbook structure are intentionally ignored
