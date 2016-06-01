namespace Rubberduck.UI.Settings
{
    /// <summary>
    /// Interaction logic for TodoSettings.xaml
    /// </summary>
    public partial class TodoSettings : ISettingsView
    {
        public TodoSettings()
        {
            InitializeComponent();
        }

        public TodoSettings(ISettingsViewModel vm)
            : this()
        {
            DataContext = vm;
        }

        public ISettingsViewModel ViewModel { get { return DataContext as ISettingsViewModel; } }
    }
}
