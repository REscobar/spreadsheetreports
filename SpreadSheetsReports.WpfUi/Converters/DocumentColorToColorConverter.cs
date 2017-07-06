namespace SpreadSheetsReports.WpfUi.Converters
{
    using System;
    using System.Globalization;
    using System.Windows.Data;
    using System.Windows.Media;

    [ValueConversion(typeof(Brush), typeof(DocumentModel.Color?))]
    public class DocumentColorToColorConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            var color = value as DocumentModel.Color?;
            if (color.HasValue)
            {
                var c = color.Value;
                return Color.FromArgb(c.Alpha, c.Red, c.Green, c.Blue);
            }

            return null;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            var color = value as Color?;
            if (color != null)
            {
                var colorValue = color.Value;
                return new DocumentModel.Color(colorValue.R, colorValue.G, colorValue.B, colorValue.A);
            }

            return null;
        }
    }
}
