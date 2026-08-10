using System.Windows;
using System.Windows.Controls;
using ExcelDiff.Models;

namespace ExcelDiff.Controls;

public partial class DiffCellControl : UserControl
{
    public static readonly DependencyProperty CellProperty = DependencyProperty.Register(
        nameof(Cell), typeof(UnifiedCell), typeof(DiffCellControl), new PropertyMetadata(null));

    public UnifiedCell? Cell
    {
        get => (UnifiedCell?)GetValue(CellProperty);
        set => SetValue(CellProperty, value);
    }

    public DiffCellControl() => InitializeComponent();
}
