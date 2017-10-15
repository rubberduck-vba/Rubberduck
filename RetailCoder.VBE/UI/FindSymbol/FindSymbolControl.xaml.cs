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
        }

        private FindSymbolViewModel ViewModel { get { return (FindSymbolViewModel)DataContext; } }

        private static readonly ICommand _goCommand = new RoutedCommand();
        public static ICommand GoCommand { get { return _goCommand; } }

        private void CommandBinding_OnExecuted(object sender, ExecutedRoutedEventArgs e)
        {
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

        private void UIElement_OnPreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter && ViewModel.CanExecute())
            {
                ViewModel.Execute();
                e.Handled = true;
            }
        }
    }
}
