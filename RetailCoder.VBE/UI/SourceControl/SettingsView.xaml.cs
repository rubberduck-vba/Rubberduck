using Ninject;

namespace Rubberduck.UI.SourceControl
{
    /// <summary>
    /// Interaction logic for SettingsView.xaml
    /// </summary>
    public partial class SettingsView
    {
        public SettingsView()
        {
            InitializeComponent();
        }

        [Inject]
        public SettingsView(SettingsViewViewModel vm) : this()
        {
            DataContext = vm;
        }
    }
}
