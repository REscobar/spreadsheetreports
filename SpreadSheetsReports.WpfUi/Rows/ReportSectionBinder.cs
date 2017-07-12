namespace SpreadSheetsReports.WpfUi.Rows
{
    using System.ComponentModel;
    using DataBinders;
    using ReportModel;

    public class ReportSectionBinder : INotifyPropertyChanged, IBinder<IRowCollectionGenerator>
    {
        private ReportSection reportSection;
        private RowCollectionBinder header;
        private ReportSectionBinder subSection;
        private RowCollectionBinder footer;

        public ReportSectionBinder()
        {
            this.Header = new RowCollectionBinder();
            this.Footer = new RowCollectionBinder();
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
                this.Header = new RowCollectionBinder();
                this.Header.ConvertFrom(reportSection.Header);
            }

            if (reportSection?.SubSection != null)
            {
                this.SubSection = new ReportSectionBinder();
                this.SubSection.ConvertFrom(reportSection.SubSection);
            }

            if (reportSection?.Footer != null)
            {
                this.Footer = new RowCollectionBinder();
                this.Footer.ConvertFrom(reportSection.Footer);
            }
        }
    }
}
