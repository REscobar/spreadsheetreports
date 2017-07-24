namespace SpreadSheetsReports.WpfUi.Rows
{
    using System.Windows;
    using System.Windows.Controls;
    using Cells;
    using Documents;

    /// <summary>
    /// Interaction logic for RowEditor.xaml
    /// </summary>
    public partial class RowEditor : UserControl
    {
        public RowEditor()
        {
            this.InitializeComponent();
        }

        private void RemoveClick(object sender, RoutedEventArgs e)
        {
            var element = sender as FrameworkElement;
            var rowBinder = element.DataContext as RowBinder;
            var rowCollection = element.Tag as RowCollectionBinder;
            rowCollection.Remove(rowBinder);
        }

        private void AddBeforeClick(object sender, RoutedEventArgs e)
        {
            var element = sender as FrameworkElement;
            var rowBinder = element.DataContext as RowBinder;
            var rowCollection = element.Tag as RowCollectionBinder;

            rowCollection.AddNewBefore(rowBinder);
        }

        private void AddAfterClick(object sender, RoutedEventArgs e)
        {
            var element = sender as FrameworkElement;
            var rowBinder = element.DataContext as RowBinder;
            var rowCollection = element.Tag as RowCollectionBinder;

            rowCollection.AddNewAfter(rowBinder);
        }
    }
}
