namespace SpreadSheetsReports.WpfUi.Rows
{
    using System.Windows;
    using System.Windows.Controls;

    /// <summary>
    /// Interaction logic for RowGroup.xaml
    /// </summary>
    public partial class ReportSectionEditor : UserControl
    {
        public ReportSectionEditor()
        {
            this.InitializeComponent();
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            var binder = this.DataContext as ReportSectionBinder;

            binder.SubSection = new ReportSectionBinder(binder.Columns);
        }

        private void AddRow_Click(object sender, RoutedEventArgs e)
        {
            var rowCollection = (sender as FrameworkElement).DataContext as RowCollectionBinder;
            rowCollection.AddNewRow();
        }
    }
}
