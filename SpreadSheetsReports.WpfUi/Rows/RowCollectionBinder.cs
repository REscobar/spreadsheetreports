namespace SpreadSheetsReports.WpfUi.Rows
{
    using System.Collections.ObjectModel;
    using System.ComponentModel;

    public class RowCollectionBinder : INotifyPropertyChanged
    {
        public RowCollectionBinder()
        {
            this.rows = new ObservableCollection<RowBinder>();
            this.AddNewRow();
            this.AddNewRow();
        }

        private readonly ObservableCollection<RowBinder> rows;

        public ObservableCollection<RowBinder> Rows
        {
            get
            {
                return rows;
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

        public event PropertyChangedEventHandler PropertyChanged;
    }
}
