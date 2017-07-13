namespace SpreadSheetsReports.WpfUi.Sheets
{
    using System;
    using System.Windows;
    using System.Windows.Controls;
    using System.Windows.Data;

    /// <summary>
    /// Interaction logic for SheetCollection.xaml
    /// </summary>
    public partial class SheetCollectionEditor : UserControl
    {
        public SheetCollectionEditor()
        {
            this.InitializeComponent();
            var view = CollectionViewSource.GetDefaultView(this.SheetsTabPanel.Items);
            view.CollectionChanged += (o, e) =>
            {
                if (e.Action == System.Collections.Specialized.NotifyCollectionChangedAction.Add)
                {
                    this.SelectTab(e.NewStartingIndex);
                }
            };
            this.SelectTab(0);
        }

        private void SelectTab(int index)
        {
            this.Dispatcher.BeginInvoke((Action)(() => this.SheetsTabPanel.SelectedIndex = index));
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
