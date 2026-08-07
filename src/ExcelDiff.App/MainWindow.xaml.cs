using System.Collections.ObjectModel;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using ExcelDiff.Models;
using ExcelDiff.Services;
using Microsoft.Win32;

namespace ExcelDiff;

public partial class MainWindow : Window
{
    private readonly IWorkbookReader _reader = new OpenXmlWorkbookReader();
    private readonly IComparisonEngine _engine = new ComparisonEngine();
    private readonly IExcelReportExporter _exporter = new ExcelReportExporter();
    private WorksheetData? _oldData;
    private WorksheetData? _newData;
    private ComparisonResult? _result;
    private CancellationTokenSource? _cancellation;

    public ObservableCollection<ColumnMapping> Mappings { get; } = [];

    public MainWindow()
    {
        InitializeComponent();
        DataContext = this;
        FieldFilter.Items.Add("All fields");
        FieldFilter.SelectedIndex = 0;
    }

    private void BrowseOld_Click(object sender, RoutedEventArgs e) => BrowseWorkbook(OldFileText, OldSheetCombo);
    private void BrowseNew_Click(object sender, RoutedEventArgs e) => BrowseWorkbook(NewFileText, NewSheetCombo);

    private void BrowseWorkbook(TextBox pathBox, ComboBox sheetCombo)
    {
        var dialog = new OpenFileDialog
        {
            Title = "Choose an Excel workbook",
            Filter = "Excel workbook (*.xlsx)|*.xlsx",
            CheckFileExists = true,
            Multiselect = false
        };
        if (dialog.ShowDialog(this) != true) return;
        try
        {
            var sheets = _reader.GetSheetNames(dialog.FileName);
            if (sheets.Count == 0) throw new InvalidDataException("This workbook has no visible worksheet records.");
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

    private async void PrepareMapping_Click(object sender, RoutedEventArgs e)
    {
        if (!TryGetInput(out var oldSheet, out var newSheet, out var oldHeader, out var newHeader)) return;
        _cancellation = new CancellationTokenSource();
        SetBusy(true, "Reading workbook headers and rows…");
        try
        {
            var oldProgress = new Progress<int>(value => BusyProgress.Value = value / 2.0);
            _oldData = await Task.Run(() => _reader.ReadSheet(OldFileText.Text, oldSheet, oldHeader, _cancellation.Token, oldProgress));
            var newProgress = new Progress<int>(value => BusyProgress.Value = 50 + value / 2.0);
            _newData = await Task.Run(() => _reader.ReadSheet(NewFileText.Text, newSheet, newHeader, _cancellation.Token, newProgress));
            BuildMappings(_oldData, _newData);
            MappingTab.IsEnabled = true;
            WorkflowTabs.SelectedItem = MappingTab;
            StatusText.Text = $"Ready to map {_oldData.Headers.Count} older fields to {_newData.Headers.Count} newer fields";
        }
        catch (OperationCanceledException)
        {
            StatusText.Text = "Reading cancelled";
        }
        catch (Exception ex)
        {
            ShowFriendlyError("The selected data could not be prepared", ex);
        }
        finally
        {
            SetBusy(false);
            _cancellation.Dispose();
            _cancellation = null;
        }
    }

    private void BuildMappings(WorksheetData oldData, WorksheetData newData)
    {
        Mappings.Clear();
        var newOptions = newData.Headers.ToArray();
        foreach (var oldHeader in oldData.Headers)
        {
            var match = newOptions.FirstOrDefault(n => NormalizeHeader(n.Name) == NormalizeHeader(oldHeader.Name));
            Mappings.Add(new ColumnMapping
            {
                OldColumn = oldHeader,
                NewColumnOptions = newOptions,
                NewColumn = match,
                IsIncluded = match is not null
            });
        }
    }

    private async void Compare_Click(object sender, RoutedEventArgs e)
    {
        if (_oldData is null || _newData is null)
        {
            MessageBox.Show(this, "Return to step 1 and prepare the workbooks again.", "Workbooks not prepared", MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }
        if (Mappings.All(m => !m.IsKey))
        {
            MessageBox.Show(this, "Select at least one Row key. A key is a field such as Employee ID or Invoice No. that uniquely identifies a row.", "Choose a row key", MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }
        if (Mappings.Any(m => m.IsKey && m.NewColumn is null))
        {
            MessageBox.Show(this, "Every selected row key must be mapped to a newer field.", "Incomplete row key", MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }

        var configuration = new ComparisonConfiguration(
            OldFileText.Text, NewFileText.Text, _oldData.SheetName, _newData.SheetName,
            _oldData.HeaderRow, _newData.HeaderRow, Mappings.ToArray(), StrictTextCheck.IsChecked == true);
        _cancellation = new CancellationTokenSource();
        SetBusy(true, "Comparing rows and values…");
        try
        {
            var progress = new Progress<int>(value => BusyProgress.Value = value);
            _result = await Task.Run(() => _engine.Compare(_oldData, _newData, configuration, _cancellation.Token, progress));
            ShowResults(_result);
            ResultsTab.IsEnabled = true;
            WorkflowTabs.SelectedItem = ResultsTab;
            StatusText.Text = $"Compared {_result.Summary.Total:N0} rows";
        }
        catch (ComparisonValidationException ex)
        {
            var details = string.Join(Environment.NewLine, ex.Issues.Take(10).Select(i => $"• {i.Message} {i.Location}"));
            if (ex.Issues.Count > 10) details += $"{Environment.NewLine}• …and {ex.Issues.Count - 10} more";
            MessageBox.Show(this, ex.Message + Environment.NewLine + Environment.NewLine + details,
                "Comparison needs attention", MessageBoxButton.OK, MessageBoxImage.Warning);
            StatusText.Text = "Fix the key fields or source data, then compare again";
        }
        catch (OperationCanceledException)
        {
            StatusText.Text = "Comparison cancelled";
        }
        catch (Exception ex)
        {
            ShowFriendlyError("The comparison could not be completed", ex);
        }
        finally
        {
            SetBusy(false);
            _cancellation.Dispose();
            _cancellation = null;
        }
    }

    private void ShowResults(ComparisonResult result)
    {
        UnchangedCount.Text = result.Summary.Unchanged.ToString("N0");
        ChangedCount.Text = result.Summary.Changed.ToString("N0");
        AddedCount.Text = result.Summary.Added.ToString("N0");
        RemovedCount.Text = result.Summary.Removed.ToString("N0");
        ProblemCount.Text = result.Summary.Problems.ToString("N0");
        ResultSubtitle.Text = $"{Path.GetFileName(result.Configuration.OldFilePath)}  →  {Path.GetFileName(result.Configuration.NewFilePath)}";

        FieldFilter.Items.Clear();
        FieldFilter.Items.Add("All fields");
        foreach (var field in result.Rows.SelectMany(r => r.Changes).Select(c => c.FieldName).Distinct(StringComparer.OrdinalIgnoreCase).OrderBy(f => f))
            FieldFilter.Items.Add(field);
        FieldFilter.SelectedIndex = 0;
        StatusFilter.SelectedIndex = 0;
        SearchText.Clear();
        ApplyFilter();
    }

    private void Filter_Changed(object sender, EventArgs e)
    {
        if (!IsLoaded || ResultsGrid is null) return;
        ApplyFilter();
    }

    private void ApplyFilter()
    {
        if (_result is null) return;
        IEnumerable<RowDifference> rows = _result.Rows;
        var status = (StatusFilter.SelectedItem as ComboBoxItem)?.Content?.ToString();
        if (!string.IsNullOrWhiteSpace(status) && status != "All statuses" && Enum.TryParse<DifferenceStatus>(status, out var parsedStatus))
            rows = rows.Where(r => r.Status == parsedStatus);
        var field = FieldFilter.SelectedItem?.ToString();
        if (!string.IsNullOrWhiteSpace(field) && field != "All fields")
            rows = rows.Where(r => r.Changes.Any(c => string.Equals(c.FieldName, field, StringComparison.OrdinalIgnoreCase)));
        var search = SearchText.Text.Trim();
        if (search.Length > 0) rows = rows.Where(r => r.SearchText.Contains(search, StringComparison.OrdinalIgnoreCase));
        ResultsGrid.ItemsSource = rows
            .OrderBy(r => r.Status switch
            {
                DifferenceStatus.Changed => 0,
                DifferenceStatus.Added => 1,
                DifferenceStatus.Removed => 2,
                _ => 3
            })
            .ThenBy(r => r.KeyDisplay, StringComparer.OrdinalIgnoreCase)
            .ToArray();
        DetailGrid.ItemsSource = null;
        SelectedKeyText.Text = "Select a row to see its values.";
    }

    private void ResultsGrid_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (ResultsGrid.SelectedItem is not RowDifference row) return;
        SelectedKeyText.Text = $"{row.StatusText}: {row.KeyDisplay}";
        if (row.Status == DifferenceStatus.Changed)
        {
            DetailGrid.ItemsSource = row.Changes;
            return;
        }
        var fields = row.OldValues.Keys.Union(row.NewValues.Keys, StringComparer.OrdinalIgnoreCase)
            .Select(field => new CellDifference(field, row.OldValues.GetValueOrDefault(field, string.Empty), row.NewValues.GetValueOrDefault(field, string.Empty)))
            .ToArray();
        DetailGrid.ItemsSource = fields;
    }

    private async void Export_Click(object sender, RoutedEventArgs e)
    {
        if (_result is null) return;
        var dialog = new SaveFileDialog
        {
            Title = "Save comparison report",
            Filter = "Excel workbook (*.xlsx)|*.xlsx",
            FileName = $"Excel comparison {DateTime.Now:yyyy-MM-dd HHmm}.xlsx",
            AddExtension = true,
            DefaultExt = ".xlsx"
        };
        if (dialog.ShowDialog(this) != true) return;
        if (IsSourceFile(dialog.FileName, _result.Configuration))
        {
            MessageBox.Show(this, "Choose a different filename. Source workbooks are never overwritten.", "Protecting source workbook", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        _cancellation = new CancellationTokenSource();
        SetBusy(true, "Creating the Excel report…");
        try
        {
            await Task.Run(() => _exporter.Export(_result, dialog.FileName, _cancellation.Token));
            StatusText.Text = $"Report saved: {dialog.FileName}";
            MessageBox.Show(this, $"The report was saved successfully.{Environment.NewLine}{dialog.FileName}", "Report ready", MessageBoxButton.OK, MessageBoxImage.Information);
        }
        catch (OperationCanceledException)
        {
            StatusText.Text = "Report export cancelled";
        }
        catch (IOException ex)
        {
            ShowFriendlyError("The report could not be saved. Close it in Excel or choose another location", ex);
        }
        catch (Exception ex)
        {
            ShowFriendlyError("The report could not be created", ex);
        }
        finally
        {
            SetBusy(false);
            _cancellation.Dispose();
            _cancellation = null;
        }
    }

    private void BackToFiles_Click(object sender, RoutedEventArgs e) => WorkflowTabs.SelectedItem = FilesTab;

    private void NewComparison_Click(object sender, RoutedEventArgs e)
    {
        _cancellation?.Cancel();
        _oldData = null;
        _newData = null;
        _result = null;
        Mappings.Clear();
        OldFileText.Clear();
        NewFileText.Clear();
        OldSheetCombo.ItemsSource = null;
        NewSheetCombo.ItemsSource = null;
        OldHeaderRowText.Text = "1";
        NewHeaderRowText.Text = "1";
        StrictTextCheck.IsChecked = false;
        ResultsGrid.ItemsSource = null;
        DetailGrid.ItemsSource = null;
        MappingTab.IsEnabled = false;
        ResultsTab.IsEnabled = false;
        WorkflowTabs.SelectedItem = FilesTab;
        StatusText.Text = "Ready for a new comparison";
    }

    private void Cancel_Click(object sender, RoutedEventArgs e)
    {
        _cancellation?.Cancel();
        StatusText.Text = "Cancelling…";
    }

    private bool TryGetInput(out string oldSheet, out string newSheet, out int oldHeader, out int newHeader)
    {
        oldSheet = OldSheetCombo.SelectedItem?.ToString() ?? string.Empty;
        newSheet = NewSheetCombo.SelectedItem?.ToString() ?? string.Empty;
        oldHeader = 0;
        newHeader = 0;
        var valid = File.Exists(OldFileText.Text) && File.Exists(NewFileText.Text)
            && oldSheet.Length > 0 && newSheet.Length > 0
            && int.TryParse(OldHeaderRowText.Text, out oldHeader) && oldHeader > 0
            && int.TryParse(NewHeaderRowText.Text, out newHeader) && newHeader > 0;
        if (!valid)
        {
            oldHeader = newHeader = 1;
            MessageBox.Show(this, "Choose both .xlsx files, select their sheets, and enter valid header row numbers.", "Complete step 1", MessageBoxButton.OK, MessageBoxImage.Information);
        }
        return valid;
    }

    private void SetBusy(bool busy, string? title = null)
    {
        BusyOverlay.Visibility = busy ? Visibility.Visible : Visibility.Collapsed;
        BusyProgress.Value = 0;
        if (title is not null) BusyTitle.Text = title;
    }

    private void ShowFriendlyError(string title, Exception exception)
    {
        var message = exception switch
        {
            UnauthorizedAccessException => "Windows denied access to the file. Check its permissions and try again.",
            InvalidDataException => exception.Message,
            NotSupportedException => exception.Message,
            _ => exception.Message
        };
        MessageBox.Show(this, message, title, MessageBoxButton.OK, MessageBoxImage.Error);
        StatusText.Text = title;
    }

    private static bool IsSourceFile(string outputPath, ComparisonConfiguration configuration) =>
        string.Equals(Path.GetFullPath(outputPath), Path.GetFullPath(configuration.OldFilePath), StringComparison.OrdinalIgnoreCase)
        || string.Equals(Path.GetFullPath(outputPath), Path.GetFullPath(configuration.NewFilePath), StringComparison.OrdinalIgnoreCase);

    private static string NormalizeHeader(string value) => value.Trim().ToUpperInvariant();
}
