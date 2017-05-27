namespace SpreadSheetsReports.WpfUi.Cells
{
    using System;
    using System.Globalization;
    using System.Windows;
    using System.Windows.Controls;
    using System.Windows.Data;
    using System.Windows.Input;
    using System.Windows.Media;
    using System.Windows.Shapes;

    /// <summary>
    /// Interaction logic for Border.xaml
    /// </summary>
    public partial class Border : UserControl
    {
        public Border()
        {
            InitializeComponent();
        }

        private void Border_MouseDown(object sender, MouseButtonEventArgs e)
        {
            var line = sender as Shape;
            var border = line.Tag as Shape;
            var color = this.colorPicker.SelectedColor;
            if ((border.Stroke as SolidColorBrush)?.Color == color)
            {
                border.Visibility = border.Visibility == Visibility.Visible ? Visibility.Hidden : Visibility.Visible;
            }
            else
            {
                border.Stroke = new SolidColorBrush(color.GetValueOrDefault());
            }
        }
    }

    [ValueConversion(typeof(Visibility),typeof(bool))]
    public class VisibilityConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            return System.Convert.ToBoolean(value) ? Visibility.Visible : Visibility.Hidden;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            return ((Visibility)value) == Visibility.Visible ? true : false;
        }
    }
}
