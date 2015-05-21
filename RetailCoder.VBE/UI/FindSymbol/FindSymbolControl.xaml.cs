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

namespace Rubberduck.UI.FindSymbol
{
    /// <summary>
    /// Interaction logic for FindSymbolControl.xaml
    /// </summary>
    public partial class FindSymbolControl : UserControl
    {
        public FindSymbolControl()
        {
            InitializeComponent();
        }

        private FindSymbolViewModel ViewModel { get { return (FindSymbolViewModel)DataContext; } }

        private static readonly ICommand _goCommand = new RoutedCommand();
        public static ICommand GoCommand { get { return _goCommand; } }

        private void CommandBinding_OnExecuted(object sender, ExecutedRoutedEventArgs e)
        {
            ViewModel.Execute();
        }

        private void CommandBinding_OnCanExecute(object sender, CanExecuteRoutedEventArgs e)
        {
            if (ViewModel == null)
            {
                return;
            }

            e.CanExecute = ViewModel.CanExecute();
            e.Handled = true;
        }
    }
}
