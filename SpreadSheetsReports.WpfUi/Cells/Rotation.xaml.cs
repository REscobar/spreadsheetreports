﻿namespace SpreadSheetsReports.WpfUi.Cells
{
    using System;
    using System.Windows;
    using System.Windows.Controls;
    using System.Windows.Input;
    using System.Windows.Shapes;
    using Xceed.Wpf.Toolkit;

    /// <summary>
    /// Interaction logic for Rotation.xaml
    /// </summary>
    public partial class Rotation : UserControl
    {
        public Rotation()
        {
            this.InitializeComponent();
        }

        private void IntegerUpDown_ValueChanged(object sender, RoutedPropertyChangedEventArgs<object> e)
        {
            var s = sender as IntegerUpDown;

            s.Value = Clamp(s.Value.Value, -90, 90);
        }

        private static T Clamp<T>(T val, T min, T max)
            where T : IComparable<T>
        {
            if (val.CompareTo(min) < 0)
            {
                return min;
            }
            else if (val.CompareTo(max) > 0)
            {
                return max;
            }
            else
            {
                return val;
            }
        }

        private void AnglePoint_MouseDown(object sender, MouseButtonEventArgs e)
        {
            var s = sender as Shape;
            var value = Convert.ToInt32(s.Tag);
            this.NUD.Value = value;
        }

        private void ContentControl_MouseMove(object sender, MouseEventArgs e)
        {
            if (e.LeftButton == MouseButtonState.Pressed)
            {
                var position = e.GetPosition(sender as IInputElement);
                var x = position.X;
                var y = 80 - position.Y;
                var angle = Math.Atan2(y, x);
                this.NUD.Value = Convert.ToInt32(angle * (180.0 / Math.PI));
            }
        }

        private void Grid_MouseUp(object sender, MouseButtonEventArgs e)
        {
            var position = e.GetPosition(sender as IInputElement);
            var x = position.X;
            var y = 80 - position.Y;
            var angle = Math.Atan2(y, x);
            this.NUD.Value = Convert.ToInt32(angle * (180.0 / Math.PI));
        }

        private void Shape_MouseMove(object sender, MouseEventArgs e)
        {
            if (e.LeftButton == MouseButtonState.Pressed)
            {
                var s = sender as Shape;
                var value = Convert.ToInt32(s.Tag);
                this.NUD.Value = value;
            }
        }
    }
}
