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
            InitializeComponent();
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            var binder = this.DataContext as ReportSectionBinder;

            binder.SubSection = new ReportSectionBinder();
        }
    }
}
