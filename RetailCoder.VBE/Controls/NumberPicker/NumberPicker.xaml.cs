using System.Windows;
using System.Windows.Controls;

// credit to http://stackoverflow.com/a/2752538
namespace Rubberduck.Controls.NumberPicker
{
    /// <summary>
    /// Interaction logic for NumberPicker.xaml
    /// </summary>
    public partial class NumberPicker
    {
        private int _numValue = 0;

        public int NumValue
        {
            get { return _numValue; }
            set
            {
                _numValue = value;
                TxtNum.Text = value.ToString();
            }
        }

        public NumberPicker()
        {
            InitializeComponent();
            TxtNum.Text = _numValue.ToString();
        }

        private void cmdUp_Click(object sender, RoutedEventArgs e)
        {
            NumValue++;
        }

        private void cmdDown_Click(object sender, RoutedEventArgs e)
        {
            NumValue--;
        }

        private void txtNum_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (TxtNum == null)
            {
                return;
            }

            if (!int.TryParse(TxtNum.Text, out _numValue))
                TxtNum.Text = _numValue.ToString();
        }
    }
}