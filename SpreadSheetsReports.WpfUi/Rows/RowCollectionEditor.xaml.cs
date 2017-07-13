namespace SpreadSheetsReports.WpfUi.Rows
{
    using System.Windows.Controls;

    /// <summary>
    /// Interaction logic for RowEditor.xaml
    /// </summary>
    public partial class RowCollectionEditor : UserControl
    {
        public RowCollectionEditor()
        {
            this.InitializeComponent();
        }

        private void NewRow_MouseUp(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            var binder = this.DataContext as RowCollectionBinder;
            binder.AddNewRow();
        }
    }
}
