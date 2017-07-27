namespace SpreadSheetsReports.WpfUi.Rows
{
    using System.Collections.Generic;
    using System.Collections.ObjectModel;
    using System.ComponentModel;
    using System.Linq;
    using Cells;
    using DataBinders;
    using DataSource;
    using ReportModel;
    using Sheets;

    public class RowBinder : INotifyPropertyChanged, IBinder<IRowGenerator>, IDataSourceBindable
    {
        private readonly ObservableCollection<CellBinder> cells;
        private readonly ObservableCollection<Column> columns;
        private readonly Dictionary<Column, CellBinder> cellMapper = new Dictionary<Column, CellBinder>();
        private int cellIndex = 0;
        private string dataMember;
        private ObservableCollection<DataSourceBinding> bindings;

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

        public string DataMember
        {
            get
            {
                return this.dataMember;
            }

            set
            {
                if (this.dataMember == value)
                {
                    return;
                }

                this.dataMember = value;

                this.NotifyPropertyChanged(nameof(this.DataMember));
            }
        }

        public ObservableCollection<DataSourceBinding> Bindings
        {
            get
            {
                return this.bindings;
            }

            set
            {
                if (this.bindings == value)
                {
                    return;
                }

                this.bindings = value;

                this.NotifyPropertyChanged(nameof(this.Bindings));
            }
        }

        public void ConvertFrom(IRowGenerator obj)
        {
            var row = obj as Row;
            if (row != null)
            {
                if (row.Bindings != null)
                {
                    this.Bindings = new ObservableCollection<DataSourceBinding>(row.Bindings.Select(b => new DataSourceBinding { Expression = b.Expression, PropertyName = b.PropertyName, Type = b.GetType().ToString() }));
                }

                var cellcount = this.Cells.Count;
                for (int i = 0; i < row.Cells.Count; i++)
                {
                    var cell = row.Cells[i];

                    CellBinder cellBinder;
                    if (cellcount <= i)
                    {
                        this.columns.Add(new Column());
                    }

                    cellBinder = this.Cells[i];
                    cellBinder.ConvertFrom(cell);
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

        private void NotifyPropertyChanged(string propertyName)
        {
            this.PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}
