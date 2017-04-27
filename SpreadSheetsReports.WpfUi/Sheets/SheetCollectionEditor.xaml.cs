namespace SpreadSheetsReports.WpfUi.Sheets
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Text;
    using System.Threading.Tasks;
    using System.Windows;
    using System.Windows.Controls;
    using System.Windows.Data;
    using System.Windows.Documents;
    using System.Windows.Input;
    using System.Windows.Media;
    using System.Windows.Media.Imaging;
    using System.Windows.Navigation;
    using System.Windows.Shapes;

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
