namespace Rubberduck.UI.Settings
{
    /// <summary>
    /// Interaction logic for UnitTestSettings.xaml
    /// </summary>
    public partial class UnitTestSettings : ISettingsView
    {
        public UnitTestSettings()
        {
            InitializeComponent();
        }

        public UnitTestSettings(ISettingsViewModel vm)
            : this()
        {
            DataContext = vm;
        }

        public ISettingsViewModel ViewModel { get { return DataContext as ISettingsViewModel; } }
    }
}
