﻿namespace SpreadSheetsReports.WpfUi.Cells
{
    using Documents;
    using System.Windows;
    using System.Windows.Controls;

    /// <summary>
    /// Interaction logic for Cell.xaml
    /// </summary>
    public partial class Cell : UserControl
    {
        public Cell()
        {
            this.InitializeComponent();
            this.MouseLeftButtonUp += this.Cell_MouseLeftButtonUp;
            this.GotFocus += this.Cell_GotFocus;
        }

        private void Cell_MouseLeftButtonUp(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            this.Focus();
        }

        private void Cell_GotFocus(object sender, System.Windows.RoutedEventArgs e)
        {
            DocumentEditor.Current.CurrentCell = this.DataContext as CellBinder;
        }

        private void FormatCells_Click(object sender, System.Windows.RoutedEventArgs e)
        {
            var cellStyleEditor = new CellStyleWindow();
            cellStyleEditor.DataContext = this.DataContext;
            cellStyleEditor.ShowDialog();
        }
    }
}
