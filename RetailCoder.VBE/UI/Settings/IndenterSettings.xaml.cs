namespace Rubberduck.UI.Settings
{
    /// <summary>
    /// Interaction logic for IndenterSettings.xaml
    /// </summary>
    public partial class IndenterSettings : ISettingsView
    {
        public IndenterSettings()
        {
            InitializeComponent();
        }
        
        public IndenterSettings(ISettingsViewModel vm)
            : this()
        {
            DataContext = vm;
        }

        public ISettingsViewModel ViewModel { get { return DataContext as ISettingsViewModel; } }
    }
}
