namespace SpreadSheetsReports.WpfUi.Rows
{
    using System.Collections.ObjectModel;
    using System.ComponentModel;

    public class RowCollectionBinder : INotifyPropertyChanged
    {
        public RowCollectionBinder()
        {
            this.rows = new ObservableCollection<RowBinder>();
            this.Rows.Add(new RowBinder());
            this.Rows.Add(new RowBinder());
            this.Rows.Add(new RowBinder());
            this.Rows.Add(new RowBinder());
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
       

        public event PropertyChangedEventHandler PropertyChanged;
    }
}
