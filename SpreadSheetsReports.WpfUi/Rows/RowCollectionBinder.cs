namespace SpreadSheetsReports.WpfUi.Rows
{
    using System.Collections.ObjectModel;
    using System.ComponentModel;
    using System.Linq;
    using DataBinders;
    using DataSource;
    using ReportModel;
    using Sheets;

    public class RowCollectionBinder : INotifyPropertyChanged, IBinder<IRowCollectionGenerator>, IDataSourceBindable
    {

        private readonly ObservableCollection<RowBinder> rows;
        private readonly ObservableCollection<Column> columns;

        private string dataMember;
        private ObservableCollection<DataSourceBinding> bindings;

        public event PropertyChangedEventHandler PropertyChanged;

        public RowCollectionBinder(ObservableCollection<Column> columns)
        {
            this.columns = columns;
            this.rows = new ObservableCollection<RowBinder>();
            this.AddNewRow();
        }

        public void Remove(RowBinder rowBinder)
        {
            this.Rows.Remove(rowBinder);
        }

        public ObservableCollection<RowBinder> Rows
        {
            get
            {
                return this.rows;
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

        public void AddNewBefore(RowBinder rowBinder)
        {
            var newRow = new RowBinder(this.columns);
            this.Rows.Insert(this.Rows.IndexOf(rowBinder), newRow);
        }

        public void AddNewAfter(RowBinder rowBinder)
        {
            var newRow = new RowBinder(this.columns);
            this.Rows.Insert(this.Rows.IndexOf(rowBinder) + 1, newRow);
        }

        private void NotifyPropertyChanged(string propertyName)
        {
            this.PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        public void AddNewRow()
        {
            this.Rows.Add(new RowBinder(this.columns));
        }

        public IRowCollectionGenerator ConvertTo()
        {
            var rows = new RowCollection();
            rows.AddRange(this.Rows.Select(r => r.ConvertTo()));
            return new RowCollectionSection
            {
                Rows = rows
            };
        }

        public void ConvertFrom(IRowCollectionGenerator obj)
        {
            var rowCollection = obj as RowCollectionSection;
            if (rowCollection != null)
            {
                if (rowCollection.Bindings != null)
                {
                    this.Bindings = new ObservableCollection<DataSourceBinding>(rowCollection.Bindings.Select(b => new DataSourceBinding { Expression = b.Expression, PropertyName = b.PropertyName, Type = b.GetType().ToString() }));
                }

                this.Rows.Clear();
                foreach (var row in rowCollection.Rows)
                {
                    var rowBinder = new RowBinder(this.columns);
                    rowBinder.ConvertFrom(row);
                    this.Rows.Add(rowBinder);
                }
            }
        }
    }
}
