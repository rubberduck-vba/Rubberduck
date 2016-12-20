namespace Rubberduck.UI.Settings
{
    /// <summary>
    /// Interaction logic for WindowSettings.xaml
    /// </summary>
    public partial class WindowSettings : ISettingsView
    {
        public WindowSettings()
        {
            InitializeComponent();
        }

        public WindowSettings(ISettingsViewModel vm)
            : this()
        {
            DataContext = vm;
        }

        public ISettingsViewModel ViewModel { get { return DataContext as ISettingsViewModel; } }
    }
}
