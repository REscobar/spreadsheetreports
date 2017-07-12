namespace SpreadSheetsReports.WpfUi.Rows
{
    using System.Collections.ObjectModel;
    using System.ComponentModel;
    using System.Linq;
    using Cells;
    using DataBinders;
    using ReportModel;

    public class RowBinder : INotifyPropertyChanged, IBinder<IRowGenerator>
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

        public event PropertyChangedEventHandler PropertyChanged;

        public ObservableCollection<CellBinder> Cells
        {
            get
            {
                return this.cells;
            }
        }

        public void ConvertFrom(IRowGenerator obj)
        {
            var row = obj as Row;
            if (row != null)
            {
                foreach (var cell in row.Cells)
                {
                    var cellBinder = new CellBinder();
                    cellBinder.ConvertFrom(cell);
                    this.Cells.Add(cellBinder);
                }
            }
        }

        public IRowGenerator ConvertTo()
        {
            return new Row
            {
                Cells = this.Cells.Select(c => c.ConvertTo()).ToList()
            };
        }
    }
}
