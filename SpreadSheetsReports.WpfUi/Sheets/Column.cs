namespace SpreadSheetsReports.WpfUi.Sheets
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel;
    using System.Linq;

    public class Column : INotifyPropertyChanged
    {
        private int index;
        private int size;

        public Column()
        {
            this.Size = 50;
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
                this.NotifyPropertyChanged(nameof(this.LetterIndex));
            }
        }

        public string LetterIndex
        {
            get
            {
                return ColumnIndexToColumnLetter(this.Index);
            }

            set
            {
                var intValue = ColumnLetterToColumnIndex(value);
                if (intValue == this.Index)
                {
                    return;
                }

                this.Index = intValue;
            }
        }

        public int Size
        {
            get
            {
                return this.size;
            }

            set
            {
                if (this.size == value)
                {
                    return;
                }

                this.size = value;
                this.NotifyPropertyChanged(nameof(this.Size));
            }
        }

        public event PropertyChangedEventHandler PropertyChanged;

        private void NotifyPropertyChanged(string propertyName)
        {
            this.PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        public static IEnumerable<Column> GetDefaultColumns()
        {
            for (int i = 0; i < 5; i++)
            {
                yield return new Column
                {
                    Index = i
                };
            }
        }

        static string ColumnIndexToColumnLetter(int colIndex)
        {
            int div = colIndex + 1;
            string colLetter = string.Empty;
            int mod = 0;

            while (div > 0)
            {
                mod = (div - 1) % 26;
                colLetter = (char)(65 + mod) + colLetter;
                div = (div - mod) / 26;
            }

            return colLetter;
        }

        static int ColumnLetterToColumnIndex(string index)
        {
            if (!string.IsNullOrEmpty(index))
            {
                return index.Aggregate(0, (s, c) => (s * 26) + c - 'A' + 1) - 1;
            }

            throw new InvalidOperationException("A column name must be specified");
        }
    }
}
