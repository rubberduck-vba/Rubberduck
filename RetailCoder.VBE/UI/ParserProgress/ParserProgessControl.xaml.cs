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

namespace Rubberduck.UI.ParserProgress
{
    /// <summary>
    /// Interaction logic for ParserProgessControl.xaml
    /// </summary>
    public partial class ParserProgessControl : UserControl
    {
        public ParserProgessControl()
        {
            InitializeComponent();
            Loaded += ParserProgessControl_Loaded;
        }

        private void ParserProgessControl_Loaded(object sender, RoutedEventArgs e)
        {
            var viewModel = DataContext as ParserProgessViewModel;
            if (viewModel == null)
            {
                return;
            }

            viewModel.Start();
        }

        public event EventHandler<ExpanderStateChangedEventArgs> ExpanderStateChanged;
        private void OnExpanderStateChanged(bool isExpanded)
        {
            var handler = ExpanderStateChanged;
            if (handler != null)
            {
                handler.Invoke(this, new ExpanderStateChangedEventArgs(isExpanded));
            }
        }

        private void Expander_OnCollapsed(object sender, RoutedEventArgs e)
        {
            OnExpanderStateChanged(false);
        }

        private void Expander_OnExpanded(object sender, RoutedEventArgs e)
        {
            OnExpanderStateChanged(true);
        }

        public class ExpanderStateChangedEventArgs : EventArgs
        {
            private readonly bool _isExpanded;

            public ExpanderStateChangedEventArgs(bool isExpanded)
            {
                _isExpanded = isExpanded;
            }

            public bool IsExpanded { get { return _isExpanded; } }
        }
    }
}
