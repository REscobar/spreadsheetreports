namespace SpreadSheetsReports.WpfUi.Cells
{
    using System.Windows;
    using System.Windows.Controls;
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
}
