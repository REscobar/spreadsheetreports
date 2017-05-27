namespace SpreadSheetsReports.WpfUi.Sheets
{
    using System.Collections.ObjectModel;
    using System.ComponentModel;

    public class SheetCollectionBinder : INotifyPropertyChanged
    {
        private readonly ObservableCollection<SheetBinder> sheets;
        private long sheetNumber = 1;

        public SheetCollectionBinder()
        {
            this.sheets = new ObservableCollection<SheetBinder>();
            this.Sheets.Add(new SheetBinder(sheetNumber++));
        }

        public event PropertyChangedEventHandler PropertyChanged;

        public ObservableCollection<SheetBinder> Sheets
        {
            get
            {
                return sheets;
            }
        }

        internal void AddSheet()
        {
            this.Sheets.Add(new SheetBinder(sheetNumber++));
        }

        internal void RemoveSheet(SheetBinder sheet)
        {
            this.Sheets.Remove(sheet);
        }
    }
}
