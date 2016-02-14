using System.Windows;

// credit to http://stackoverflow.com/a/2752538
namespace Rubberduck.UI.Controls
{
    /// <summary>
    /// Interaction logic for NumberPicker.xaml
    /// </summary>
    public partial class NumberPicker
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
    }
}