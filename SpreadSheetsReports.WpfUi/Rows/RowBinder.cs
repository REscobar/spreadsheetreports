namespace SpreadSheetsReports.WpfUi.Rows
{
    using System.Collections.Generic;
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
        private readonly Dictionary<Column, CellBinder> cellMapper = new Dictionary<Column, CellBinder>();
        private int cellIndex = 0;

        public event PropertyChangedEventHandler PropertyChanged;

        public RowBinder(ObservableCollection<Column> columns)
        {
            this.columns = columns;
            this.cells = new ObservableCollection<CellBinder>();
            foreach (var column in columns)
            {
                this.AddCell(column);
            }
            this.Columns.CollectionChanged += this.Columns_CollectionChanged;
        }

        private void Columns_CollectionChanged(object sender, System.Collections.Specialized.NotifyCollectionChangedEventArgs e)
        {
            if (e.Action == System.Collections.Specialized.NotifyCollectionChangedAction.Add)
            {
                foreach (Column column in e.NewItems)
                {
                    this.AddCell(column);
                }
            }
            else if (e.Action == System.Collections.Specialized.NotifyCollectionChangedAction.Remove)
            {
                foreach (Column column in e.OldItems)
                {
                    this.RemoveCell(column);
                }
            }
        }

        private void RemoveCell(Column column)
        {
            CellBinder cell;
            this.cellMapper.TryGetValue(column, out cell);
            if (cell != null)
            {
                this.cellMapper.Remove(column);
                this.cells.Remove(cell);
            }
        }

        public void AddCell(Column column)
        {
            var cell = new CellBinder(this.cellIndex++, column);
            this.cells.Add(cell);
            this.cellMapper[column] = cell;
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
                return this.columns;
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
