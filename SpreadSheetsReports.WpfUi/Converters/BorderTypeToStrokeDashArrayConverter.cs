namespace SpreadSheetsReports.WpfUi.Converters
{
    using System;
    using System.Globalization;
    using System.Windows.Data;
    using System.Windows.Media;

    [ValueConversion(typeof(DocumentModel.BorderType), typeof(DoubleCollection))]
    public class BorderTypeToStrokeDashArrayConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            var borderType = value as DocumentModel.BorderType?;
            if (borderType.HasValue)
            {
                var collection = new DoubleCollection();
                switch (borderType.Value)
                {
                    case DocumentModel.BorderType.None:
                        break;
                    case DocumentModel.BorderType.Thin:
                    case DocumentModel.BorderType.Medium:
                    case DocumentModel.BorderType.Thick:
                    case DocumentModel.BorderType.Hair:
                    case DocumentModel.BorderType.Double:
                        break;
                    case DocumentModel.BorderType.Dashed:
                    case DocumentModel.BorderType.MediumDashed:
                        collection.Add(5);
                        collection.Add(2);
                        break;
                    case DocumentModel.BorderType.Dotted:
                        collection.Add(2);
                        collection.Add(1);
                        break;
                    case DocumentModel.BorderType.DashDot:
                    case DocumentModel.BorderType.MediumDashDot:
                    case DocumentModel.BorderType.SlantedDashDot:
                        collection.Add(5);
                        collection.Add(2);
                        collection.Add(2);
                        collection.Add(2);
                        break;
                    case DocumentModel.BorderType.DashDotDot:
                    case DocumentModel.BorderType.MediumDashDotDot:
                        collection.Add(5);
                        collection.Add(2);
                        collection.Add(2);
                        collection.Add(2);
                        collection.Add(2);
                        collection.Add(2);
                        break;
                    default:
                        break;
                }

                return collection;
            }

            return null;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            return null;
        }
    }
}
