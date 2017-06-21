namespace SpreadSheetsReports.WpfUi.Rows
{
    using System.Collections.ObjectModel;
    using System.ComponentModel;
    using SpreadSheetsReports.WpfUi.Cells;

    public class RowBinder : INotifyPropertyChanged
    {
        private readonly ObservableCollection<CellBinder> cells;
        public RowBinder()
        {
            this.cells = new ObservableCollection<CellBinder>();
            this.cells.Add(new CellBinder());
            this.cells.Add(new CellBinder());
            this.cells.Add(new CellBinder());
            this.cells.Add(new CellBinder());
        }

        public ObservableCollection<CellBinder> Cells
        {
            get
            {
                return cells;
            }
        }

        public event PropertyChangedEventHandler PropertyChanged;
    }
}
