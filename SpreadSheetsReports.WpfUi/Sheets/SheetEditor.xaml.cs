namespace SpreadSheetsReports.WpfUi.Sheets
{
    using System.Collections.ObjectModel;
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

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            var columns = ((sender as FrameworkElement).Tag as ObservableCollection<Column>);
            columns.Add(new Column());
        }
    }
}
