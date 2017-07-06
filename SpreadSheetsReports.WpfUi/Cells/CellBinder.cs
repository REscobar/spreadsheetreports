namespace SpreadSheetsReports.WpfUi.Cells
{
    using System.ComponentModel;
    using DocumentModel;
    using SpreadSheetsReports.ReportModel;

    public class CellBinder : INotifyPropertyChanged
    {
        private CellStyle style;
        private string className;
        private object value;
        private CellType type;

        public CellBinder()
        {
            this.Value = "Sample";
        }

        /// <inheritdoc/>
        public event PropertyChangedEventHandler PropertyChanged;

        /// <summary>
        /// Gets or sets the <see cref="CellStyle"/>
        /// </summary>
        public CellStyle Style
        {
            get
            {
                return this.style;
            }

            set
            {
                if (this.style == value)
                {
                    return;
                }

                this.style = value;
                this.NotifyPropertyChanged(nameof(this.Style));
            }
        }

        /// <summary>
        /// Gets or sets style class name
        /// </summary>
        public string ClassName
        {
            get
            {
                return this.className;
            }

            set
            {
                if (this.className == value)
                {
                    return;
                }

                this.className = value;
                this.NotifyPropertyChanged(nameof(this.ClassName));
            }
        }

        /// <summary>
        /// Gets or sets the cell value
        /// </summary>
        public object Value
        {
            get
            {
                return this.value;
            }

            set
            {
                if (this.value == value)
                {
                    return;
                }

                this.value = value;
                this.NotifyPropertyChanged(nameof(this.Value));
            }
        }

        /// <summary>
        /// Gets or sets the <see cref="CellType"/>
        /// </summary>
        public CellType Type
        {
            get
            {
                return this.type;
            }

            set
            {
                if (this.type == value)
                {
                    return;
                }

                this.type = value;
                this.NotifyPropertyChanged(nameof(this.Type));
            }
        }

        public void EnsureStyleIsCreated()
        {
            if (this.Style == null)
            {
                this.Style = new CellStyle();
            }
        }

        private void NotifyPropertyChanged(string propertyName)
        {
            this.PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

    }
}
