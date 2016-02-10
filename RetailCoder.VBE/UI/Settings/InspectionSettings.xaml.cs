namespace Rubberduck.UI.Settings
{
    /// <summary>
    /// Interaction logic for InspectionSettings.xaml
    /// </summary>
    public partial class InspectionSettings : ISettingsView
    {
        public InspectionSettings()
        {
            InitializeComponent();
        }

        public InspectionSettings(InspectionSettingsViewModel vm) : this()
        {
            DataContext = vm;
        }
    }
}
