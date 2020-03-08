using System.Windows.Controls;

namespace Rubberduck.UI.Settings
{
    /// <summary>
    /// Interaction logic for AutoCompleteSettings.xaml
    /// </summary>
    public partial class AutoCompleteSettings : UserControl, ISettingsView
    {
        public AutoCompleteSettings()
        {
            InitializeComponent();
        }

        public AutoCompleteSettings(ISettingsViewModel vm) : this()
        {
            DataContext = vm;
        }

        public ISettingsViewModel ViewModel => DataContext as ISettingsViewModel;
    }
}
