namespace Rubberduck.UI.Settings
{
    /// <summary>
    /// Interaction logic for GeneralSettings.xaml
    /// </summary>
    public partial class GeneralSettings : ISettingsView
    {
        public GeneralSettings()
        {
            InitializeComponent();
        }

        public GeneralSettings(GeneralSettingsViewModel vm) : this()
        {
            DataContext = vm;
        }
    }
}
