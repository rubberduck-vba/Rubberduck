using System.Windows.Controls;

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

        private void FindSymbolControl_Loaded(object sender, System.Windows.RoutedEventArgs e)
        {
            SearchComboBox.Focus();
        }
    }
}