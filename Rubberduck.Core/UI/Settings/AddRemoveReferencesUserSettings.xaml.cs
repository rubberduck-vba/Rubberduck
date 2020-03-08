namespace Rubberduck.UI.Settings
{
    /// <summary>
    /// Interaction logic for AddRemoveReferencesUserSettings.xaml
    /// </summary>
    public partial class AddRemoveReferencesUserSettings : ISettingsView
    {
        public AddRemoveReferencesUserSettings()
        {
            InitializeComponent();
        }

        public AddRemoveReferencesUserSettings(ISettingsViewModel viewModel) : this()
        {
            DataContext = viewModel;
        }

        public ISettingsViewModel ViewModel => DataContext as ISettingsViewModel;
    }
}
