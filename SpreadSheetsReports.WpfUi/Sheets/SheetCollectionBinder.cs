namespace SpreadSheetsReports.WpfUi.Sheets
{
    using Cells;
    using System.Collections.ObjectModel;
    using System.ComponentModel;
    using System.Windows;

    public class SheetCollectionBinder : INotifyPropertyChanged
    {
        private readonly ObservableCollection<SheetBinder> sheets;
        private long sheetNumber = 1;

        public SheetCollectionBinder()
        {
            this.sheets = new ObservableCollection<SheetBinder>();
            this.AddSheet();
        }

        /// <inheritdoc/>
        public event PropertyChangedEventHandler PropertyChanged;

        public ObservableCollection<SheetBinder> Sheets
        {
            get
            {
                return sheets;
            }
        }

        private void NotifyPropertyChanged(string propertyName)
        {
            this.PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        internal void AddSheet()
        {
            var sheet = new SheetBinder(this.sheetNumber++);
            this.Sheets.Add(sheet);
        }

        internal void RemoveSheet(SheetBinder sheet)
        {
            this.Sheets.Remove(sheet);
        }
    }
}
