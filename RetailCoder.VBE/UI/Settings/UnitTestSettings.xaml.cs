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

        public UnitTestSettings(UnitTestSettingsViewModel vm)
            : this()
        {
            DataContext = vm;
        }
    }
}
