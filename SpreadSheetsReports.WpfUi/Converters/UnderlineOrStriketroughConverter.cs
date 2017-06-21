namespace SpreadSheetsReports.WpfUi.Converters
{
    using System;
    using System.Globalization;
    using System.Windows;
    using System.Windows.Data;


    public class UnderlineOrStriketroughConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            var style = value as DocumentModel.FontStyle;
            if (style != null)
            {
                return style.Underline == DocumentModel.UnderLineStyle.Single ? TextDecorations.Underline : style.IsStrikeout ? TextDecorations.Strikethrough : null;
            }

            return null;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}
