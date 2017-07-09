namespace SpreadSheetsReports.WpfUi.Converters
{
    using System;
    using System.Collections.Generic;
    using System.Globalization;
    using System.Linq;
    using System.Text;
    using System.Threading.Tasks;
    using System.Windows.Data;

    [ValueConversion(typeof(DocumentModel.BorderType), typeof(double))]
    public class BorderTypeToAngleConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            var borderType = value as DocumentModel.BorderType?;
            if (borderType.HasValue)
            {
                switch (borderType.Value)
                {
                    case DocumentModel.BorderType.SlantedDashDot:
                        return -45d;
                    case DocumentModel.BorderType.None:
                    case DocumentModel.BorderType.Thin:
                    case DocumentModel.BorderType.Medium:
                    case DocumentModel.BorderType.Thick:
                    case DocumentModel.BorderType.Hair:
                    case DocumentModel.BorderType.Double:
                    case DocumentModel.BorderType.Dashed:
                    case DocumentModel.BorderType.MediumDashed:
                    case DocumentModel.BorderType.Dotted:
                    case DocumentModel.BorderType.DashDot:
                    case DocumentModel.BorderType.MediumDashDot:
                    case DocumentModel.BorderType.DashDotDot:
                    case DocumentModel.BorderType.MediumDashDotDot:
                    default:
                        break;
                }
            }

            return 0d;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}
