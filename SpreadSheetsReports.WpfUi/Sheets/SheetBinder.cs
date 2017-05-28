namespace SpreadSheetsReports.WpfUi.Sheets
{
    using System.ComponentModel;
    using SpreadSheetsReports.WpfUi.Rows;
    using Cells;

    public class SheetBinder : INotifyPropertyChanged
    {
        private ReportSectionBinder content;
        private string name;
        private CellBinder current;

        public event PropertyChangedEventHandler PropertyChanged;

        public SheetBinder()
            : this(0)
        {
        }

        public SheetBinder(long sheetNumber)
        {
            this.Content = new ReportSectionBinder();
            this.Name = $"Sheet{sheetNumber}";
        }

        public ReportSectionBinder Content
        {
            get
            {
                return this.content;
            }

            set
            {
                if (this.content == value)
                {
                    return;
                }

                this.content = value;
                this.NotifyPropertyChanged(nameof(this.Content));
            }
        }

        public CellBinder CurrentCell
        {
            get
            {
                return this.current;
            }

            set
            {
                if (this.current == value)
                {
                    return;
                }

                this.current = value;
                this.NotifyPropertyChanged(nameof(this.CurrentCell));
            }
        }

        public string Name
        {
            get
            {
                return this.name;
            }

            set
            {
                if (this.name == value)
                {
                    return;
                }

                this.name = value;
                this.NotifyPropertyChanged(nameof(this.Name));
            }
        }

        private void NotifyPropertyChanged(string propertyName)
        {
            this.PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

    }
}
