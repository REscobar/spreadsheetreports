namespace SpreadSheetsReports.WpfUi.Sheets
{
    using System.Windows;
    using System.Windows.Controls;

    /// <summary>
    /// Interaction logic for Sheet.xaml
    /// </summary>
    public partial class SheetEditor : UserControl
    {
        public SheetEditor()
        {
            this.InitializeComponent();
        }

        private void Resize_Click(object sender, RoutedEventArgs e)
        {
            var item = sender as FrameworkElement;
            var resizer = new ColumnResizerWindow();
            resizer.DataContext = item.DataContext;
            resizer.ShowDialog();
        }
    }
}
