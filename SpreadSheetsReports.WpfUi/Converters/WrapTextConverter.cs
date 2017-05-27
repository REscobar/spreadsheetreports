namespace SpreadSheetsReports.WpfUi.Converters
{
    using System;
    using System.Globalization;
    using System.Windows;
    using System.Windows.Data;
    using System.Windows.Media;

    public class WrapTextConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            var wrapText = value as bool?;
            if (wrapText.HasValue)
            {
                return wrapText.Value ? TextWrapping.Wrap : TextWrapping.NoWrap;
            }
            return null;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}
