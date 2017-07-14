namespace SpreadSheetsReports.WpfUi.Cells
{
    using System.ComponentModel;
    using DataBinders;
    using DocumentModel;
    using ReportModel;
    using Sheets;

    public class CellBinder : INotifyPropertyChanged, IBinder<ICellGenerator>
    {
        private CellStyle style;
        private string className;
        private object value;
        private CellType type;

        private int index;
        private int width;

        public CellBinder()
        {
            this.Value = "Sample";
        }

        public CellBinder(int index)
            : this()
        {
            this.index = index;
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

        public int Index
        {
            get
            {
                return this.index;
            }

            set
            {
                if (this.index == value)
                {
                    return;
                }

                this.index = value;
                this.NotifyPropertyChanged(nameof(this.Index));
            }
        }

        public int Width
        {
            get
            {
                return this.width;
            }

            set
            {
                if (this.width == value)
                {
                    return;
                }

                this.width = value;

                this.NotifyPropertyChanged(nameof(this.Width));
            }
        }

        public void AssignColumn(Column column)
        {
            this.Width = column.Size;
            column.PropertyChanged += (s, e) =>
            {
                if (e.PropertyName == nameof(Column.Size))
                {
                    this.Width = column.Size;
                }
            };
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

        public ICellGenerator ConvertTo()
        {
            return new ReportModel.Cell
            {
                Style = this.Style,
                Type = this.Type,
                Value = this.Value
            };
        }

        public void ConvertFrom(ICellGenerator obj)
        {
            var cell = obj as ReportModel.Cell;
            this.Style = cell.Style;
            this.Type = cell.Type;
            this.Value = cell.Value;
        }
    }
}
