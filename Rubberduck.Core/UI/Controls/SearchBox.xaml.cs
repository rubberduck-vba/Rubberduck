using NLog;
using Rubberduck.UI.Command;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

namespace Rubberduck.UI.Controls
{
    /// <summary>
    /// Interaction logic for SearchBox.xaml
    /// </summary>
    public partial class SearchBox : UserControl
    {
        public static readonly DependencyProperty TextProperty =
            DependencyProperty.Register(nameof(Text), typeof(string), typeof(SearchBox), new UIPropertyMetadata(default(string), PropertyChangedCallback));
        public static readonly DependencyProperty HintProperty =
            DependencyProperty.Register(nameof(Hint), typeof(string), typeof(SearchBox), new UIPropertyMetadata(default(string), PropertyChangedCallback));

        private static void PropertyChangedCallback(DependencyObject source, DependencyPropertyChangedEventArgs e)
        {
            if (source is SearchBox control)
            {
                var newValue = (string)e.NewValue;
                switch (e.Property.Name)
                {
                    case nameof(Text):
                        control.Text = newValue;
                        break;
                    case nameof(Hint):
                        control.Hint = newValue;
                        break;
                }
            }
        }
        
        public string Text
        {
            get => (string)GetValue(TextProperty);
            set
            {
                var old = GetValue(TextProperty);
                SetValue(TextProperty, value);
                OnPropertyChanged(new DependencyPropertyChangedEventArgs(TextProperty, old, value));
            }
        }
        public string Hint
        {
            get => (string)GetValue(HintProperty);
            set
            {
                var old = GetValue(HintProperty);
                SetValue(HintProperty, value);
                OnPropertyChanged(new DependencyPropertyChangedEventArgs(HintProperty, old, value));
            }
        }

        public ICommand ClearSearchCommand => new DelegateCommand(LogManager.GetCurrentClassLogger(), (arg) => Text = string.Empty);

        public SearchBox()
        {
            // design instance!
            Text = string.Empty;
            Hint = "Search";
            Width = 300;
            Height = 25;
            Background = SystemColors.WindowBrush;
            // not so much design instance
            InitializeComponent();
        }
    }
}
