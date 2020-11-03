using System.Windows;
using System.Windows.Controls.Primitives;
using Rubberduck.AddRemoveReferences;

namespace Rubberduck.UI.AddRemoveReferences
{
    /// <summary>
    /// Interaction logic for AddRemoveReferencesWindow.xaml
    /// </summary>
    public partial class AddRemoveReferencesWindow
    {
        public AddRemoveReferencesWindow()
        {
            InitializeComponent();
        }

        private AddRemoveReferencesViewModel ViewModel => DataContext as AddRemoveReferencesViewModel;

        private void ListView_SynchronizeCurrentSelection_OnGotFocus(object sender, RoutedEventArgs e)
        {
            UpdateCurrentSelection((Selector)sender);

            var cs = ViewModel.CurrentSelection;
            Description.Text = cs?.Description ?? string.Empty;
            Version.Text = cs?.Version ?? string.Empty;
            LocaleName.Text = cs?.LocaleName ?? string.Empty;
            FullPath.Text = cs?.FullPath ?? string.Empty;
        }

        private void UpdateCurrentSelection(Selector sender)
        {
            var selectedReferenceModel = (ReferenceModel)sender.SelectedItem;
            ViewModel.CurrentSelection = selectedReferenceModel;
        }
    }
}
