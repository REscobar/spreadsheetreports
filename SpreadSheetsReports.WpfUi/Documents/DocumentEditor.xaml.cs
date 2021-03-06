﻿namespace SpreadSheetsReports.WpfUi.Documents
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
    using SpreadSheetsReports.WpfUi.Cells;

    /// <summary>
    /// Interaction logic for DocumentEditor.xaml
    /// </summary>
    public partial class DocumentEditor : UserControl
    {
        public static DocumentEditor Current { get; set; }

        public DocumentEditor()
        {
            this.InitializeComponent();
            Current = this;
        }

        public CellBinder CurrentCell
        {
            get { return (CellBinder)this.GetValue(CurrentCellProperty); }
            set { this.SetValue(CurrentCellProperty, value); }
        }

        // Using a DependencyProperty as the backing store for CurrentCell.  This enables animation, styling, binding, etc...
        public static readonly DependencyProperty CurrentCellProperty =
            DependencyProperty.Register("CurrentCell", typeof(CellBinder), typeof(DocumentEditor), new PropertyMetadata(null));
    }
}
