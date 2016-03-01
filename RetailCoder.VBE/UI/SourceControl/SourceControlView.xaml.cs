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
            var pwd = new SecureString();
            foreach (var c in PasswordBox.Password)
            {
                pwd.AppendChar(c);
            }

            var vm = (SourceControlViewViewModel)DataContext;
            vm.CreateProviderWithCredentials(new SecureCredentials(UserNameBox.Text, pwd));
        }
    }
}
