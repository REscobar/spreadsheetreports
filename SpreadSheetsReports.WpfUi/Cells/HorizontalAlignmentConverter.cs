namespace SpreadSheetsReports.WpfUi.Cells
{
    using System;
    using System.Globalization;
    using System.Windows;
    using System.Windows.Data;
    using System.Windows.Media;
    
    public class HorizontalAlignmentConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            var align = value as DocumentModel.HorizontalAlignment?;
            if (align.HasValue)
            {
                switch (align.Value)
                {
                    case DocumentModel.HorizontalAlignment.General:
                    case DocumentModel.HorizontalAlignment.Left:
                        return HorizontalAlignment.Left;
                    case DocumentModel.HorizontalAlignment.Center:
                        return HorizontalAlignment.Center;
                    case DocumentModel.HorizontalAlignment.Right:
                        return HorizontalAlignment.Right;
                    case DocumentModel.HorizontalAlignment.Fill:
                    case DocumentModel.HorizontalAlignment.Justify:
                    case DocumentModel.HorizontalAlignment.CenterAcrossSelection:
                    case DocumentModel.HorizontalAlignment.Distributed:
                        return HorizontalAlignment.Stretch;
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
