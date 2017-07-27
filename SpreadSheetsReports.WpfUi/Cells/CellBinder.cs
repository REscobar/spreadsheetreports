namespace SpreadSheetsReports.WpfUi.Cells
{
    using System.Collections.ObjectModel;
    using System.ComponentModel;
    using DataBinders;
    using DataSource;
    using DocumentModel;
    using ReportModel;
    using Sheets;
    using System.Linq;

    public class CellBinder : INotifyPropertyChanged, IBinder<ICellGenerator>, IDataSourceBindable
    {
        private CellStyle style;
        private string className;
        private object value;
        private CellType type;

        private int index;
        private int width;
        private string dataMember;
        private ObservableCollection<DataSourceBinding> bindings;

        public CellBinder()
        {
            this.Value = this.Value ?? string.Empty;
            this.Bindings = this.Bindings ?? new ObservableCollection<DataSourceBinding>();
        }

        public CellBinder(int index, Column column)
            : this()
        {
            this.index = index;
            this.AssignColumn(column);
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

        private void AssignColumn(Column column)
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
                Value = this.Value,
                Bindings = new PropertyBindingCollection(this.Bindings.Select(b => b.ConvertTo()))
            };
        }

        public void ConvertFrom(ICellGenerator obj)
        {
            var cell = obj as ReportModel.Cell;
            if (cell != null)
            {

                if (cell.Bindings != null)
                {
                    this.Bindings = new ObservableCollection<DataSourceBinding>(cell.Bindings.Select(b => new DataSourceBinding { Expression = b.Expression, PropertyName = b.PropertyName, Type = b.GetType().Name }));
                }

                this.Style = cell.Style;
                this.Type = cell.Type;
                this.Value = cell.Value;
            }
        }
    }
}
