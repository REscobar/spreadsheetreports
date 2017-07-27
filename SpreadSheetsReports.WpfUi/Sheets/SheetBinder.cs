namespace SpreadSheetsReports.WpfUi.Sheets
{
    using System;
    using System.Collections.ObjectModel;
    using System.ComponentModel;
    using Cells;
    using DataBinders;
    using DataSource;
    using ReportModel;
    using Rows;
    using System.Linq;

    public class SheetBinder : INotifyPropertyChanged, IBinder<Sheet>, IDataSourceBindable
    {
        private readonly ObservableCollection<Column> columns;
        private ReportSectionBinder content;
        private string name;
        private CellBinder current;
        private string dataMember;
        private ObservableCollection<DataSourceBinding> bindings;

        public event PropertyChangedEventHandler PropertyChanged;

        public SheetBinder()
            : this(0)
        {
        }

        public SheetBinder(long sheetNumber)
        {
            this.columns = new ObservableCollection<Column>(Column.GetDefaultColumns());
            this.Content = new ReportSectionBinder(this.Columns);
            this.Bindings = new ObservableCollection<DataSourceBinding>();
            this.Name = $"Sheet{sheetNumber}";
            this.columns.CollectionChanged += this.Columns_CollectionChanged;
        }

        private void Columns_CollectionChanged(object sender, System.Collections.Specialized.NotifyCollectionChangedEventArgs e)
        {
            for (int i = 0; i < this.Columns.Count; i++)
            {
                this.Columns[i].Index = i;
            }
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
                if (obj.Bindings != null)
                {
                    this.Bindings = new ObservableCollection<DataSourceBinding>(obj.Bindings.Select(b => new DataSourceBinding { Expression = b.Expression, PropertyName = b.PropertyName, Type = b.GetType().ToString() }));
                }

                this.Content.ConvertFrom(obj.Content);
            }
        }
    }
}
