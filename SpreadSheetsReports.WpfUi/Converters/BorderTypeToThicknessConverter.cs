namespace SpreadSheetsReports.WpfUi.Converters
{
    using System;
    using System.Globalization;
    using System.Windows.Data;
    using DocumentModel;

    [ValueConversion(typeof(BorderType), typeof(double))]
    class BorderTypeToThicknessConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            var borderType = value as BorderType?;
            if (borderType.HasValue)
            {
                switch (borderType.Value)
                {
                    case BorderType.None:
                        return 0D;
                    case BorderType.Thin:
                    case BorderType.Dashed:
                    case BorderType.DashDot:
                    case BorderType.DashDotDot:
                    case BorderType.Hair:
                    case BorderType.Double:
                        return 1D;
                    case BorderType.Medium:
                    case BorderType.Dotted:
                    case BorderType.MediumDashed:
                    case BorderType.MediumDashDot:
                    case BorderType.MediumDashDotDot:
                    case BorderType.SlantedDashDot:
                        return 2D;
                    case BorderType.Thick:
                        return 4D;
                }
            }

            return 0D;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}
