namespace SpreadSheetsReports.WpfUi.Sheets
{
    using System.Collections.Generic;
    using System.Collections.ObjectModel;
    using System.ComponentModel;
    using System.Linq;
    using DataBinders;
    using ReportModel;

    public class SheetCollectionBinder : INotifyPropertyChanged, IBinder<IEnumerable<ISheetGenerator>>
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
                return this.sheets;
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

        public IEnumerable<ISheetGenerator> ConvertTo()
        {
            return this.Sheets.Select(s => s.ConvertTo());
        }

        public void ConvertFrom(IEnumerable<ISheetGenerator> obj)
        {
            var sheets = obj == null ? Enumerable.Empty<Sheet>() : obj.OfType<Sheet>();

            if (sheets != null)
            {
                foreach (var sheet in sheets)
                {
                    var sheetBinder = new SheetBinder();
                    sheetBinder.ConvertFrom(sheet);

                    this.Sheets.Add(sheetBinder);
                }
            }
        }
    }
}
