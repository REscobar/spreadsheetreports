namespace SpreadSheetsReports.WpfUi.Cells
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel;
    using System.Linq;
    using System.Text;
    using System.Threading.Tasks;
    using SpreadSheetsReports.DocumentModel;
    using SpreadSheetsReports.ReportModel;

    public class CellBinder : INotifyPropertyChanged
    {
        private CellStyle? style;
        private string className;
        private object value;
        private CellType type;

        public CellBinder()
        {
            this.Value = "TestingTestingTestingTesting";
            this.Style = new CellStyle
            {
                HorizontalAlignment = HorizontalAlignment.Distributed,
                VerticalAlignment = VerticalAlignment.Justify,
                BackgroundColor = new Color(alpha:127, green:212,blue:212),
                Rotation = 12,
                ShrinkToFit = true,
                WrapText = true,
                FontStyle = new FontStyle
                {
                    FontName = "Comic Sans MS",
                    Color = new Color(red:255),
                    IsBold = true,
                    IsItalic = true,
                    IsStrikeout = true,
                    Underline = UnderLineStyle.Single,
                    Size = 12,
                    ScriptStyle = FontScriptStyle.Superscript
                }
            };
        }

        /// <inheritdoc/>
        public event PropertyChangedEventHandler PropertyChanged;

        /// <summary>
        /// Gets or sets the <see cref="CellStyle"/>
        /// </summary>
        public CellStyle? Style
        {
            get
            {
                return this.style;
            }

            set
            {
                if (this.style.Equals(value))
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

        private void NotifyPropertyChanged(string propertyName)
        {
            this.PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

    }
}
