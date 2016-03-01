using System.Security;
using Rubberduck.SourceControl;

namespace Rubberduck.UI.SourceControl
{
    /// <summary>
    /// Interaction logic for SourceControlPanel.xaml
    /// </summary>
    public partial class SourceControlView
    {
        public SourceControlView()
        {
            InitializeComponent();
        }

        private void Login_Click(object sender, System.Windows.RoutedEventArgs e)
        {
            var vm = (SourceControlViewViewModel) DataContext;

            var pwd = new SecureString();
            foreach (var c in PasswordBox.Password)
            {
                pwd.AppendChar(c);
            }

            vm.Provider = vm.ProviderFactory.CreateProvider(vm.VBE.ActiveVBProject, vm.Provider.CurrentRepository,
                new SecureCredentials(UserNameBox.Text, pwd), vm.WrapperFactory);
        }
    }
}
