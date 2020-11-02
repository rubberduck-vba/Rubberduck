namespace Rubberduck.UI.Settings
{
    /// <summary>
    /// Interaction logic for AddRemoveReferencesUserSettings.xaml
    /// </summary>
    public partial class IgnoredProjectsSettingsView : ISettingsView
    {
        public IgnoredProjectsSettingsView()
        {
            InitializeComponent();
        }

        public IgnoredProjectsSettingsView(ISettingsViewModel viewModel) : this()
        {
            DataContext = viewModel;
        }

        public ISettingsViewModel ViewModel => DataContext as ISettingsViewModel;
    }
}
