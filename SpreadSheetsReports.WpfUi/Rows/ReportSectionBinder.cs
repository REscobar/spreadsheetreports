namespace SpreadSheetsReports.WpfUi.Rows
{
    using System.Collections.ObjectModel;
    using System.ComponentModel;
    using DataBinders;
    using DataSource;
    using ReportModel;
    using Sheets;

    public class ReportSectionBinder : INotifyPropertyChanged, IBinder<IRowCollectionGenerator>, IDataSourceBindable
    {
        private readonly ObservableCollection<Column> columns;
        private ReportSection reportSection;
        private RowCollectionBinder header;
        private ReportSectionBinder subSection;
        private RowCollectionBinder footer;
        private string dataMember;
        private ObservableCollection<DataSourceBinding> bindings;

        public ReportSectionBinder()
        {
        }

        public ReportSectionBinder(ObservableCollection<Column> columns)
        {
            this.columns = columns;

            this.Header = new RowCollectionBinder(this.Columns);
            this.Footer = new RowCollectionBinder(this.Columns);
        }

        /// <inheritdoc/>
        public event PropertyChangedEventHandler PropertyChanged;

        public ReportSection ReportSection
        {
            get
            {
                return this.reportSection;
            }

            set
            {
                if (this.reportSection == value)
                {
                    return;
                }

                this.reportSection = value;

                this.NotifyPropertyChanged(nameof(this.ReportSection));
            }
        }

        public RowCollectionBinder Header
        {
            get
            {
                return this.header;
            }

            set
            {
                if (this.header == value)
                {
                    return;
                }

                this.header = value;
                this.NotifyPropertyChanged(nameof(this.Header));
            }
        }

        public ReportSectionBinder SubSection
        {
            get
            {
                return this.subSection;
            }

            set
            {
                if (this.subSection == value)
                {
                    return;
                }

                this.subSection = value;
                this.NotifyPropertyChanged(nameof(this.SubSection));
            }
        }

        public RowCollectionBinder Footer
        {
            get
            {
                return this.footer;
            }

            set
            {
                if (this.footer == value)
                {
                    return;
                }

                this.footer = value;
                this.NotifyPropertyChanged(nameof(this.Footer));
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

        public IRowCollectionGenerator ConvertTo()
        {
            return new ReportSection
            {
                Header = this.Header?.ConvertTo(),
                SubSection = this.SubSection?.ConvertTo(),
                Footer = this.Footer?.ConvertTo()
            };
        }

        public void ConvertFrom(IRowCollectionGenerator obj)
        {
            var reportSection = obj as ReportSection;
            if (reportSection?.Header != null)
            {
                this.Header = new RowCollectionBinder(this.columns);
                this.Header.ConvertFrom(reportSection.Header);
            }

            if (reportSection?.SubSection != null)
            {
                this.SubSection = new ReportSectionBinder(this.columns);
                this.SubSection.ConvertFrom(reportSection.SubSection);
            }

            if (reportSection?.Footer != null)
            {
                this.Footer = new RowCollectionBinder(this.columns);
                this.Footer.ConvertFrom(reportSection.Footer);
            }
        }
    }
}
