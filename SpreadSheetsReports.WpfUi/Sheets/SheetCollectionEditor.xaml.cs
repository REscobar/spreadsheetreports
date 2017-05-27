namespace SpreadSheetsReports.WpfUi.Sheets
{
    using System.Windows;
    using System.Windows.Controls;

    /// <summary>
    /// Interaction logic for SheetCollection.xaml
    /// </summary>
    public partial class SheetCollectionEditor : UserControl
    {
        public SheetCollectionEditor()
        {
            InitializeComponent();
        }

        private void button_Click(object sender, RoutedEventArgs e)
        {
            var binder = this.DataContext as SheetCollectionBinder;
            binder.AddSheet();
        }

        private void MenuItem_Click(object sender, RoutedEventArgs e)
        {
            var sheet = (sender as Control).DataContext as SheetBinder;
            var binder = this.DataContext as SheetCollectionBinder;
            binder.RemoveSheet(sheet);
        }
    }
}
