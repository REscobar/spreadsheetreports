namespace SpreadSheetsReports.WpfUi.Converters
{
    using System;
    using System.Globalization;
    using System.Windows.Data;
    using DocumentModel;

    class BorderTypeToVisibilityConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            var borderType = value as BorderType?;
            if (borderType.HasValue)
            {
                switch (borderType.Value)
                {
                    case BorderType.None:
                        return false;
                    case BorderType.Thin:
                    case BorderType.Medium:
                    case BorderType.Dashed:
                    case BorderType.Dotted:
                    case BorderType.Thick:
                    case BorderType.Double:
                    case BorderType.Hair:
                    case BorderType.MediumDashed:
                    case BorderType.DashDot:
                    case BorderType.MediumDashDot:
                    case BorderType.DashDotDot:
                    case BorderType.MediumDashDotDot:
                    case BorderType.SlantedDashDot:
                        return true;
                }
            }

            return false;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            var hasBorder = value as bool?;
            if (hasBorder.HasValue)
            {
                return BorderType.Thin;
            }
            return BorderType.None;
        }
    }
}
