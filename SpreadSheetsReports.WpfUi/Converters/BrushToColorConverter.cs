namespace SpreadSheetsReports.WpfUi.Converters
{
    using System;
    using System.Globalization;
    using System.Windows.Data;
    using System.Windows.Media;


    [ValueConversion(typeof(Brush), typeof(DocumentModel.Color?))]
    public class BrushToColorConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            var color = value as DocumentModel.Color?;
            if (color.HasValue)
            {
                var c = color.Value;
                return new SolidColorBrush(Color.FromArgb(c.Alpha, c.Red, c.Green, c.Blue));
            }

            return null;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            var brush = value as SolidColorBrush;
            if (brush != null)
            {
                return new DocumentModel.Color(brush.Color.R, brush.Color.G, brush.Color.B, brush.Color.A);
            }

            return null;
        }
    }
}
