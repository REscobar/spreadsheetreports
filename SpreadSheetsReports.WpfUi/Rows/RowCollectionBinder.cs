namespace SpreadSheetsReports.WpfUi.Rows
{
    using System.Collections.ObjectModel;
    using System.ComponentModel;
    using System.Linq;
    using DataBinders;
    using ReportModel;
    using Sheets;

    public class RowCollectionBinder : INotifyPropertyChanged, IBinder<IRowCollectionGenerator>
    {

        public RowCollectionBinder(ObservableCollection<Column> columns)
        {
            this.columns = columns;
            this.rows = new ObservableCollection<RowBinder>();
            this.AddNewRow();
            this.AddNewRow();
        }

        public event PropertyChangedEventHandler PropertyChanged;

        private readonly ObservableCollection<RowBinder> rows;
        private readonly ObservableCollection<Column> columns;

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
                return columns;
            }
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
            if (obj != null)
            {
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
