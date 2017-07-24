namespace SpreadSheetsReports.WpfUi.Cells
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
    using System.Windows.Shapes;

    /// <summary>
    /// Interaction logic for CellStyleWindow.xaml
    /// </summary>
    public partial class CellStyleWindow : Window
    {
        public CellStyleWindow()
        {
            this.InitializeComponent();
        }

        private void Window_DataContextChanged(object sender, DependencyPropertyChangedEventArgs e)
        {
            var cell = e.NewValue as CellBinder;
            if (cell != null)
            {
                cell.EnsureStyleIsCreated();
            }
        }
    }
}
