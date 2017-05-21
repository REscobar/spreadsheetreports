﻿namespace SpreadSheetsReports.WpfUi.Cells
{
    using System;
    using System.Collections.Generic;
    using System.Globalization;
    using System.Linq;
    using System.Text;
    using System.Threading.Tasks;
    using System.Windows;
    using System.Windows.Controls;
    using System.Windows.Data;
    using System.Windows.Documents;
    using System.Windows.Input;
    using System.Windows.Media;
    using System.Windows.Media.Imaging;
    using System.Windows.Navigation;
    using System.Windows.Shapes;


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
                return new DocumentModel.Color(brush.Color.A, brush.Color.R, brush.Color.G, brush.Color.B);
            }

            return null;
        }
    }

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
                        collection.Add(2);
                        break;
                    case DocumentModel.BorderType.Dashed:

                        break;
                    case DocumentModel.BorderType.Dotted:
                        collection.Add(2);
                        collection.Add(1);
                        break;
                    case DocumentModel.BorderType.Double:
                        break;
                    case DocumentModel.BorderType.MediumDashed:
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
            var brush = value as SolidColorBrush;
            if (brush != null)
            {
                return new DocumentModel.Color(brush.Color.A, brush.Color.R, brush.Color.G, brush.Color.B);
            }

            return null;
        }
    }
}
