using System.Windows.Controls;

namespace Rubberduck.UI.AddRemoveReferences
{
    /// <summary>
    /// Interaction logic for AddRemoveReferencesWindow.xaml
    /// </summary>
    public partial class AddRemoveReferencesWindow : UserControl
    {
        public AddRemoveReferencesWindow()
        {
            InitializeComponent();
        }

        private AddRemoveReferencesViewModel ViewModel => DataContext as AddRemoveReferencesViewModel;
    }
}
