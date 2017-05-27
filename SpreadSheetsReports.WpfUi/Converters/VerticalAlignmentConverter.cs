namespace SpreadSheetsReports.WpfUi.Converters
{
    using System;
    using System.Globalization;
    using System.Windows;
    using System.Windows.Data;

    public class VerticalAlignmentConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            var align = value as DocumentModel.VerticalAlignment?;
            if (align.HasValue)
            {
                switch (align.Value)
                {
                    case DocumentModel.VerticalAlignment.Bottom:
                        return VerticalAlignment.Bottom;
                    case DocumentModel.VerticalAlignment.Top:
                        return VerticalAlignment.Top;
                    case DocumentModel.VerticalAlignment.Center:
                        return VerticalAlignment.Center;
                    case DocumentModel.VerticalAlignment.Justify:
                    case DocumentModel.VerticalAlignment.Distributed:
                    case DocumentModel.VerticalAlignment.JustifyDistributed:
                        return VerticalAlignment.Stretch;
                    default:
                        break;
                }
            }
            return null;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}
