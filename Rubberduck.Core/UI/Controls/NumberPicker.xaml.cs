using System;
using System.ComponentModel;
using System.Windows;
using System.Windows.Controls;
using Rubberduck.Settings;

// credit to http://stackoverflow.com/a/2752538
namespace Rubberduck.UI.Controls
{
    /// <summary>
    /// Interaction logic for NumberPicker.xaml
    /// </summary>
    public partial class NumberPicker : UserControl, IDataErrorInfo
    {
        public static readonly DependencyProperty NumValueProperty =
            DependencyProperty.Register(nameof(NumValue), typeof(int), typeof(NumberPicker), new UIPropertyMetadata(default(int), PropertyChangedCallback));
        public static readonly DependencyProperty MinNumberProperty =
            DependencyProperty.Register(nameof(MinNumber), typeof(int), typeof(NumberPicker), new UIPropertyMetadata(default(int), PropertyChangedCallback));
        public static readonly DependencyProperty MaxNumberProperty =
            DependencyProperty.Register(nameof(MaxNumber), typeof(int), typeof(NumberPicker), new UIPropertyMetadata(default(int), PropertyChangedCallback));

        private static void PropertyChangedCallback(DependencyObject source, DependencyPropertyChangedEventArgs args)
        {
            if (source is NumberPicker control)
            {
                var newValue = (int) args.NewValue;
                switch (args.Property.Name)
                {
                    case "NumValue":
                        control.NumValue = newValue;
                        break;
                    case "MinNumber":
                        control.MinNumber = newValue;
                        break;
                    case "MaxNumber":
                        control.MaxNumber = newValue;
                        break;
                }
            }
        }

        public int NumValue
        {
            get => (int)GetValue(NumValueProperty);
            set
            {
                var old = GetValue(MinNumberProperty);
                SetValue(NumValueProperty, value);
                OnPropertyChanged(new DependencyPropertyChangedEventArgs(NumValueProperty, old, value));
            }
        }

        public int MinNumber
        {
            get => (int)GetValue(MinNumberProperty);
            set
            {
                var old = GetValue(MinNumberProperty);
                SetValue(MinNumberProperty, value);
                OnPropertyChanged(new DependencyPropertyChangedEventArgs(MinNumberProperty, old, value));              
            }
        }

        public int MaxNumber
        {
            get => (int)GetValue(MaxNumberProperty);
            set
            {
                var old = GetValue(MaxNumberProperty);
                SetValue(MaxNumberProperty, value);
                OnPropertyChanged(new DependencyPropertyChangedEventArgs(MaxNumberProperty, old, value));
            }
        }

        public NumberPicker()
        {
            InitializeComponent();
        }

        private void cmdUp_Click(object sender, RoutedEventArgs e)
        {
            NumValue++;
        }

        private void cmdDown_Click(object sender, RoutedEventArgs e)
        {
            NumValue--;
        }

        public string this[string columnName]
        {
            get
            {
                if (columnName != "NumValue")
                {
                    return string.Empty;
                }

                if (NumValue < MinNumber || NumValue > MaxNumber)
                {
                    return "Invalid Selection";
                }

                return string.Empty;
            }
        }

        public string Error { get; private set; }
    }
}
