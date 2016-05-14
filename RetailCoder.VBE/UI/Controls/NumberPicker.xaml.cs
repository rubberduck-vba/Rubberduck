using System.ComponentModel;
using System.Windows;

// credit to http://stackoverflow.com/a/2752538
namespace Rubberduck.UI.Controls
{
    /// <summary>
    /// Interaction logic for NumberPicker.xaml
    /// </summary>
    public partial class NumberPicker : IDataErrorInfo
    {
        public static readonly DependencyProperty NumValueProperty =
            DependencyProperty.Register("NumValue", typeof(int), typeof(NumberPicker), new UIPropertyMetadata(null));

        public int NumValue
        {
            get
            {
                return (int)GetValue(NumValueProperty);
            }
            set
            {
                SetValue(NumValueProperty, value);
                OnPropertyChanged(new DependencyPropertyChangedEventArgs(NumValueProperty, NumValue - 1, NumValue));
            }
        }

        private int _minNumber = int.MinValue;
        public int MinNumber
        {
            get { return _minNumber; }
            set { _minNumber = value; }
        }

        private int _maxNumber = int.MaxValue;
        public int MaxNumber
        {
            get { return _maxNumber; }
            set { _maxNumber = value; }
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
