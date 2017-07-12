namespace SpreadSheetsReports.WpfUi.Rows
{
    using System.Collections.ObjectModel;
    using System.ComponentModel;
    using System.Linq;
    using DataBinders;
    using ReportModel;

    public class RowCollectionBinder : INotifyPropertyChanged, IBinder<IRowCollectionGenerator>
    {
        public RowCollectionBinder()
        {
            this.rows = new ObservableCollection<RowBinder>();
            this.AddNewRow();
            this.AddNewRow();
        }

        public event PropertyChangedEventHandler PropertyChanged;

        private readonly ObservableCollection<RowBinder> rows;

        public ObservableCollection<RowBinder> Rows
        {
            get
            {
                return this.rows;
            }
        }

        private void NotifyPropertyChanged(string propertyName)
        {
            this.PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        public void AddNewRow()
        {
            this.Rows.Add(new RowBinder());
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
                    var rowBinder = new RowBinder();
                    rowBinder.ConvertFrom(row);
                    this.Rows.Add(rowBinder);
                }
            }
        }
    }
}
