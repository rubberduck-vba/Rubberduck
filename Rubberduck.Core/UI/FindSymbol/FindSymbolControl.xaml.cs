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
            ViewModel?.Execute();
            e.Handled = true;
        }

        private void CommandBinding_OnCanExecute(object sender, CanExecuteRoutedEventArgs e)
        {
            e.CanExecute = ViewModel?.CanExecute() ?? false;
            e.Handled = true;
        }

        private void FindSymbolControl_Loaded(object sender, System.Windows.RoutedEventArgs e)
        {
            searchComboBox.Focus();
        }

        // doing this navigates on arrow-up/down, which isn't expected 
        //private void searchComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e) => ViewModel?.Execute();
    }
}