using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Threading;
using ExcelDiff.Controls;
using ExcelDiff.Models;
using ExcelDiff.Services;
using Microsoft.Win32;

namespace ExcelDiff;

public partial class MainWindow : Window
{
    private readonly IWorkbookReader _reader = new WorkbookReader();
    private readonly IWorksheetDiffEngine _engine = new WorksheetDiffEngine();
    private WorksheetDiffResult? _result;
    private IReadOnlyList<UnifiedDiffRow> _differenceRows = [];
    private int _differenceIndex = -1;
    private CancellationTokenSource? _cancellation;

    public MainWindow()
    {
        InitializeComponent();
        Loaded += MainWindow_Loaded;
        UpdateNavigation();
    }

    private async void MainWindow_Loaded(object sender, RoutedEventArgs e)
    {
        var arguments = Environment.GetCommandLineArgs().Skip(1).ToArray();
        var demoIndex = Array.FindIndex(arguments, value => string.Equals(value, "--demo", StringComparison.OrdinalIgnoreCase));
        if (demoIndex < 0) return;
        if (demoIndex + 2 >= arguments.Length)
        {
            MessageBox.Show(this, "Demo mode requires paths to the older and newer Excel workbooks.", "Invalid demo command", MessageBoxButton.OK, MessageBoxImage.Error);
            return;
        }
        var screenshotIndex = Array.FindIndex(arguments, value => string.Equals(value, "--screenshot", StringComparison.OrdinalIgnoreCase));
        var screenshotPath = screenshotIndex >= 0 && screenshotIndex + 1 < arguments.Length ? arguments[screenshotIndex + 1] : null;
        await LoadAndCompareAsync(arguments[demoIndex + 1], arguments[demoIndex + 2], null, null, screenshotPath);
    }

    private void BrowseOld_Click(object sender, RoutedEventArgs e) => BrowseWorkbook(OldFileText, OldSheetCombo);
    private void BrowseNew_Click(object sender, RoutedEventArgs e) => BrowseWorkbook(NewFileText, NewSheetCombo);

    private void BrowseWorkbook(TextBox pathBox, ComboBox sheetCombo)
    {
        var dialog = new OpenFileDialog
        {
            Title = "Choose an Excel workbook",
            Filter = "Excel workbooks (*.xlsx;*.xls)|*.xlsx;*.xls|Modern workbook (*.xlsx)|*.xlsx|Legacy workbook (*.xls)|*.xls",
            CheckFileExists = true,
            Multiselect = false
        };
        if (dialog.ShowDialog(this) != true) return;
        try
        {
            var sheets = _reader.GetSheetNames(dialog.FileName);
            if (sheets.Count == 0) throw new InvalidDataException("This workbook has no worksheets.");
            pathBox.Text = dialog.FileName;
            sheetCombo.ItemsSource = sheets;
            sheetCombo.SelectedIndex = 0;
            StatusText.Text = $"Loaded {Path.GetFileName(dialog.FileName)}";
        }
        catch (Exception ex)
        {
            ShowFriendlyError("The workbook could not be opened", ex);
        }
    }

    private async void Compare_Click(object sender, RoutedEventArgs e)
    {
        var oldPath = OldFileText.Text;
        var newPath = NewFileText.Text;
        var oldSheet = OldSheetCombo.SelectedItem?.ToString();
        var newSheet = NewSheetCombo.SelectedItem?.ToString();
        if (!File.Exists(oldPath) || !File.Exists(newPath) || string.IsNullOrWhiteSpace(oldSheet) || string.IsNullOrWhiteSpace(newSheet))
        {
            MessageBox.Show(this, "Choose both workbooks and one worksheet from each.", "Choose worksheets", MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }
        await LoadAndCompareAsync(oldPath, newPath, oldSheet, newSheet, null);
    }

    private async Task LoadAndCompareAsync(string oldPath, string newPath, string? oldSheet, string? newSheet, string? screenshotPath)
    {
        _cancellation = new CancellationTokenSource();
        SetBusy(true, "Reading worksheet values…");
        try
        {
            var oldSheets = _reader.GetSheetNames(oldPath);
            var newSheets = _reader.GetSheetNames(newPath);
            oldSheet ??= oldSheets.FirstOrDefault();
            newSheet ??= newSheets.FirstOrDefault();
            if (oldSheet is null || newSheet is null) throw new InvalidDataException("Both workbooks must contain a worksheet.");

            OldFileText.Text = oldPath;
            NewFileText.Text = newPath;
            OldSheetCombo.ItemsSource = oldSheets;
            NewSheetCombo.ItemsSource = newSheets;
            OldSheetCombo.SelectedItem = oldSheet;
            NewSheetCombo.SelectedItem = newSheet;

            var token = _cancellation.Token;
            var oldProgress = new Progress<int>(value => BusyProgress.Value = value * 0.35);
            var oldGrid = await Task.Run(() => _reader.ReadGrid(oldPath, oldSheet, token, oldProgress));
            var newProgress = new Progress<int>(value => BusyProgress.Value = 35 + value * 0.35);
            var newGrid = await Task.Run(() => _reader.ReadGrid(newPath, newSheet, token, newProgress));
            SetBusy(true, "Aligning rows and columns…");
            var compareProgress = new Progress<int>(value => BusyProgress.Value = 70 + value * 0.30);
            _result = await Task.Run(() => _engine.Compare(oldGrid, newGrid, token, compareProgress));

            ShowResult(_result, oldPath, newPath, oldSheet, newSheet, oldGrid, newGrid);
            if (!string.IsNullOrWhiteSpace(screenshotPath))
            {
                SetBusy(false);
                await Dispatcher.InvokeAsync(() => { }, DispatcherPriority.ApplicationIdle);
                CaptureApplicationContent(screenshotPath);
                StatusText.Text += " • Screenshot saved";
            }
        }
        catch (OperationCanceledException)
        {
            StatusText.Text = "Comparison cancelled";
        }
        catch (Exception ex)
        {
            ShowFriendlyError("The worksheets could not be compared", ex);
        }
        finally
        {
            SetBusy(false);
            _cancellation.Dispose();
            _cancellation = null;
        }
    }

    private void ShowResult(WorksheetDiffResult result, string oldPath, string newPath, string oldSheet, string newSheet, WorksheetGrid oldGrid, WorksheetGrid newGrid)
    {
        BuildGridColumns(result.Columns);
        DiffGrid.ItemsSource = result.Rows;
        _differenceRows = result.Rows.Where(row => row.IsDifference).ToArray();
        _differenceIndex = -1;
        ResultsTab.IsEnabled = true;
        WorkflowTabs.SelectedItem = ResultsTab;
        var range = UsedRange(oldGrid, newGrid);
        ResultSubtitle.Text = $"{Path.GetFileName(oldPath)} [{oldSheet}]  →  {Path.GetFileName(newPath)} [{newSheet}]  •  {range}";
        StatusText.Text = result.DifferenceRowCount == 0
            ? "No differences found"
            : $"{result.DifferenceRowCount:N0} changed row{(result.DifferenceRowCount == 1 ? "" : "s")} • {result.ChangedCellCount:N0} changed cell{(result.ChangedCellCount == 1 ? "" : "s")}";
        UpdateNavigation();
        if (_differenceRows.Count > 0) Navigate(1);
    }

    private void BuildGridColumns(IReadOnlyList<UnifiedColumn> columns)
    {
        DiffGrid.Columns.Clear();
        DiffGrid.Columns.Add(new DataGridTextColumn
        {
            Header = "Row",
            Binding = new Binding(nameof(UnifiedDiffRow.RowLabel)),
            Width = 145,
            IsReadOnly = true
        });
        for (var index = 0; index < columns.Count; index++)
        {
            var factory = new FrameworkElementFactory(typeof(DiffCellControl));
            factory.SetBinding(DiffCellControl.CellProperty, new Binding($"Cells[{index}]"));
            DiffGrid.Columns.Add(new DataGridTemplateColumn
            {
                Header = columns[index].Label,
                CellTemplate = new DataTemplate { VisualTree = factory },
                Width = new DataGridLength(180),
                MinWidth = 90,
                IsReadOnly = true
            });
        }
    }

    private void PreviousDifference_Click(object sender, RoutedEventArgs e) => Navigate(-1);
    private void NextDifference_Click(object sender, RoutedEventArgs e) => Navigate(1);

    private void Navigate(int direction)
    {
        if (_differenceRows.Count == 0) return;
        _differenceIndex = Math.Clamp(_differenceIndex + direction, 0, _differenceRows.Count - 1);
        var row = _differenceRows[_differenceIndex];
        DiffGrid.SelectedItem = row;
        DiffGrid.ScrollIntoView(row);
        DiffGrid.Focus();
        UpdateNavigation();
    }

    private void UpdateNavigation()
    {
        PreviousDifferenceButton.IsEnabled = _differenceIndex > 0;
        NextDifferenceButton.IsEnabled = _differenceRows.Count > 0 && _differenceIndex < _differenceRows.Count - 1;
        DifferencePositionText.Text = _differenceRows.Count == 0
            ? "No changed rows"
            : $"Changed row {_differenceIndex + 1:N0} of {_differenceRows.Count:N0}";
    }

    private void NewComparison_Click(object sender, RoutedEventArgs e)
    {
        _cancellation?.Cancel();
        _result = null;
        _differenceRows = [];
        _differenceIndex = -1;
        OldFileText.Clear();
        NewFileText.Clear();
        OldSheetCombo.ItemsSource = null;
        NewSheetCombo.ItemsSource = null;
        DiffGrid.ItemsSource = null;
        DiffGrid.Columns.Clear();
        ResultsTab.IsEnabled = false;
        WorkflowTabs.SelectedItem = FilesTab;
        StatusText.Text = "Ready for a new comparison";
        UpdateNavigation();
    }

    private void Cancel_Click(object sender, RoutedEventArgs e)
    {
        _cancellation?.Cancel();
        StatusText.Text = "Cancelling…";
    }

    private void SetBusy(bool busy, string? title = null)
    {
        BusyOverlay.Visibility = busy ? Visibility.Visible : Visibility.Collapsed;
        if (busy) BusyProgress.Value = 0;
        if (title is not null) BusyTitle.Text = title;
    }

    private void ShowFriendlyError(string title, Exception exception)
    {
        var message = exception switch
        {
            UnauthorizedAccessException => "Windows denied access to the file. Check its permissions and try again.",
            InvalidDataException or NotSupportedException => exception.Message,
            _ => exception.Message
        };
        MessageBox.Show(this, message, title, MessageBoxButton.OK, MessageBoxImage.Error);
        StatusText.Text = title;
    }

    private static string UsedRange(WorksheetGrid oldGrid, WorksheetGrid newGrid)
    {
        var maxColumn = Math.Max(oldGrid.MaxColumn, newGrid.MaxColumn);
        var maxRow = Math.Max(oldGrid.MaxRow, newGrid.MaxRow);
        return maxColumn == 0 || maxRow == 0 ? "Empty used range" : $"A1:{ExcelColumnName(maxColumn - 1)}{maxRow}";
    }

    private static string ExcelColumnName(int zeroBasedIndex)
    {
        var value = zeroBasedIndex + 1;
        var name = string.Empty;
        while (value > 0)
        {
            value--;
            name = (char)('A' + value % 26) + name;
            value /= 26;
        }
        return name;
    }

    private void CaptureApplicationContent(string outputPath)
    {
        if (Content is not Visual visual || Content is not FrameworkElement contentElement)
            throw new InvalidOperationException("The application content is not ready to capture.");
        contentElement.UpdateLayout();
        var dpi = VisualTreeHelper.GetDpi(contentElement);
        var bitmap = new RenderTargetBitmap(
            Math.Max(1, (int)Math.Ceiling(contentElement.ActualWidth * dpi.DpiScaleX)),
            Math.Max(1, (int)Math.Ceiling(contentElement.ActualHeight * dpi.DpiScaleY)),
            96 * dpi.DpiScaleX, 96 * dpi.DpiScaleY, PixelFormats.Pbgra32);
        bitmap.Render(visual);
        var directory = Path.GetDirectoryName(Path.GetFullPath(outputPath));
        if (!string.IsNullOrWhiteSpace(directory)) Directory.CreateDirectory(directory);
        var encoder = new PngBitmapEncoder();
        encoder.Frames.Add(BitmapFrame.Create(bitmap));
        using var stream = new FileStream(outputPath, FileMode.Create, FileAccess.Write, FileShare.None);
        encoder.Save(stream);
    }
}
