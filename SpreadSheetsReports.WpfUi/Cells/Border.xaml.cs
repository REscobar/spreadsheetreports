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
            this.InitializeComponent();
        }

        private void Border_MouseDown(object sender, MouseButtonEventArgs e)
        {
            this.ApplyBorder(sender);
        }

        private void ApplyBorder(object sender)
        {
            var line = sender as FrameworkElement;
            var border = line.Tag as Shape;
            var color = this.colorPicker.SelectedColor;

            PropertyDescriptorCollection properties = TypeDescriptor.GetProperties(border.Tag);

            // Sets an PropertyDescriptor to the specific property.
            PropertyDescriptor myProperty = properties.Find("Type", false);

            object borderType = myProperty.GetValue(border.Tag);
            object newBorderType = (this.listBox.SelectedItem as FrameworkElement).Tag;

            if ((border.Stroke as SolidColorBrush)?.Color == color && borderType.Equals(newBorderType))
            {
                borderType = DocumentModel.BorderType.None;
            }
            else
            {
                borderType = newBorderType;
            }


            myProperty.SetValue(border.Tag, borderType);

            var descriptor = DependencyPropertyDescriptor.FromName(
                   "Stroke",
                   border.GetType(),
                   border.GetType());

            // now you can set property value with
            descriptor.SetValue(border, new SolidColorBrush(color.GetValueOrDefault()));
        }

        private void BorderVisibilityClick(object sender, RoutedEventArgs e)
        {
            this.ApplyBorder(sender);
        }
    }
}
