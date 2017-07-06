namespace SpreadSheetsReports.WpfUi.Cells
{
    using System.ComponentModel;
    using System.Windows;
    using System.Windows.Controls;
    using System.Windows.Controls.Primitives;
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
                var descriptor = DependencyPropertyDescriptor.FromName(
                    "Stroke",
                    border.GetType(),
                    border.GetType());

                // now you can set property value with
                descriptor.SetValue(border, new SolidColorBrush(color.GetValueOrDefault()));

                // Creates a new collection and assign it the properties for button1.
                PropertyDescriptorCollection properties = TypeDescriptor.GetProperties(border.Tag);

                // Sets an PropertyDescriptor to the specific property.
                System.ComponentModel.PropertyDescriptor myProperty = properties.Find("Type", false);

                myProperty.SetValue(border.Tag, (listBox.SelectedItem as FrameworkElement).Tag);

                // also, you can use the dependency property itself
                //var property = descriptor.DependencyProperty;
                //dependencyObject.SetValue(property, value);
                //border.Stroke = new SolidColorBrush(color.GetValueOrDefault());
            }
        }

        private void BorderVisibilityClick(object sender, RoutedEventArgs e)
        {
            var button = sender as ToggleButton;
            button.IsChecked = !button.IsChecked;
        }

        private void Right_IsVisibleChanged(object sender, DependencyPropertyChangedEventArgs e)
        {

        }
    }
}
