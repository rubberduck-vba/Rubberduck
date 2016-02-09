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

        public TodoSettings(TodoSettingsViewModel vm) : this()
        {
            DataContext = vm;
        }
    }
}
