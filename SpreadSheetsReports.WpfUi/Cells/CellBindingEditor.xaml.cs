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

namespace SpreadSheetsReports.WpfUi.Cells
{
    /// <summary>
    /// Interaction logic for CellBindingEditor.xaml
    /// </summary>
    public partial class CellBindingEditor : UserControl
    {
        public CellBindingEditor()
        {
            InitializeComponent();
        }

        internal void Ok()
        {
            (this.DataContext as CellBindingsViewModel).Copy();
        }

        internal void SetBinder(CellBinder binder)
        {
            this.DataContext = new CellBindingsViewModel(binder);
        }
    }
}
