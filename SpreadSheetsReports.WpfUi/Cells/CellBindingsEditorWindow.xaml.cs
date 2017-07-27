namespace SpreadSheetsReports.WpfUi.Cells
{
    using System.Windows;

    /// <summary>
    /// Interaction logic for CellBindingsEditor.xaml
    /// </summary>
    public partial class CellBindingsEditorWindow : Window
    {
        public CellBindingsEditorWindow(CellBinder binder)
        {
            InitializeComponent();
            this.Editor.SetBinder(binder);
        }

        private void Ok_Click(object sender, RoutedEventArgs e)
        {
            this.Editor.Ok();
            this.DialogResult = true;
        }

        private void Cancel_Click(object sender, RoutedEventArgs e)
        {
            this.DialogResult = false;
        }
    }
}
