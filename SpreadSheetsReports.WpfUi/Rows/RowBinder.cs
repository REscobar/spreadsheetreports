namespace SpreadSheetsReports.WpfUi.Rows
{
    using System.Collections.ObjectModel;
    using System.ComponentModel;
    using System.Linq;
    using Cells;
    using DataBinders;
    using ReportModel;
    using Sheets;

    public class RowBinder : INotifyPropertyChanged, IBinder<IRowGenerator>
    {
        private readonly ObservableCollection<CellBinder> cells;
        private readonly ObservableCollection<Column> columns;
        private int cellIndex = 0;

        public RowBinder(ObservableCollection<Column> columns)
        {
            this.columns = columns;
            this.cells = new ObservableCollection<CellBinder>();
            this.AddCell();
        }

        public event PropertyChangedEventHandler PropertyChanged;

        public void AddCell()
        {
            var cell = new CellBinder(this.cellIndex++);
            cell.AssignColumn(this.columns[cell.Index]);
            this.cells.Add(cell);
        }

        public ObservableCollection<CellBinder> Cells
        {
            get
            {
                return this.cells;
            }
        }

        public ObservableCollection<Column> Columns
        {
            get
            {
                return columns;
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
