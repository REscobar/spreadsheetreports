namespace SpreadSheetsReports.WpfUi.Sheets
{
    using System.ComponentModel;
    using SpreadSheetsReports.WpfUi.Rows;
    using Cells;
    using DataBinders;
    using ReportModel;
    using System;
    using System.Collections.ObjectModel;

    public class SheetBinder : INotifyPropertyChanged, IBinder<Sheet>
    {
        private ReportSectionBinder content;
        private string name;
        private CellBinder current;

        private readonly ObservableCollection<Column> columns;

        public event PropertyChangedEventHandler PropertyChanged;

        public SheetBinder()
            : this(0)
        {
        }

        public SheetBinder(long sheetNumber)
        {
            this.columns = new ObservableCollection<Column>(Column.GetDefaultColumns());
            this.Content = new ReportSectionBinder(this.Columns);
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

        public Sheet ConvertTo()
        {
            return new Sheet
            {
                Name = this.Name,
                Content = this.Content.ConvertTo()
            };
        }

        public void ConvertFrom(Sheet obj)
        {
            this.Name = obj.Name;
            if (obj.Content != null)
            {
                this.Content = new ReportSectionBinder(this.Columns);
                this.Content.ConvertFrom(obj.Content);
            }
        }
    }
}
