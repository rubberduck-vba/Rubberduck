using Ninject;

namespace Rubberduck.UI.SourceControl
{
    /// <summary>
    /// Interaction logic for SettingsView.xaml
    /// </summary>
    public partial class SettingsView : IControlView
    {
        public SettingsView()
        {
            InitializeComponent();
        }

        [Inject]
        public SettingsView(IControlViewModel vm) : this()
        {
            DataContext = vm;
        }

        public IControlViewModel ViewModel { get { return (IControlViewModel)DataContext; } }
    }
}
