namespace SpreadSheetsReports.WpfUi.Sheets
{
    using System;
    using System.Collections.Generic;
    using System.Collections.ObjectModel;
    using System.ComponentModel;
    using System.Linq;
    using System.Text;
    using System.Threading.Tasks;

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
