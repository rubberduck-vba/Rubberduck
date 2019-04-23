using NLog;
using Rubberduck.UI.Command;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

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
                    case "Text":
                        control.Text = newValue;
                        break;
                    case "Hint":
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

        public ICommand ClearSearchCommand { get => new DelegateCommand(LogManager.GetCurrentClassLogger(), (arg) => Text = ""); }

        public SearchBox()
        {
            // design instance!
            Text = "";
            Hint = "Search";
            Width = 300;
            Height = 25;
            Background = SystemColors.WindowBrush;
            // not so much design instance
            InitializeComponent();
        }
    }
}
