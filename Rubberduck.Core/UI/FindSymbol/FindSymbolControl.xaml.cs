using System.Windows.Controls;
using System.Windows.Input;

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
            Loaded += FindSymbolControl_Loaded;
        }

        private FindSymbolViewModel ViewModel => (FindSymbolViewModel)DataContext;

        public static ICommand GoCommand { get; } = new RoutedCommand();

        private void CommandBinding_OnExecuted(object sender, ExecutedRoutedEventArgs e)
        {
            if (ViewModel == null)
            {
                return;
            }
            
            ViewModel.Execute();
            e.Handled = true;
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

        private void FindSymbolControl_Loaded(object sender, System.Windows.RoutedEventArgs e)
        {
            searchComboBox.Focus();
        }
    }
}