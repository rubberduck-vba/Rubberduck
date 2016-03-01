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
            var vm = (SourceControlViewViewModel)DataContext;
            vm.CreateProviderWithCredentials(new SecureCredentials(UserNameBox.Text, PasswordBox.SecurePassword));

            PasswordBox.Password = string.Empty;
        }
    }
}
