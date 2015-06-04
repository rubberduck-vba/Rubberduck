using System.Windows.Forms;

namespace Rubberduck.UI.SourceControl
{
    public partial class SettingsControl : UserControl, ISettingsView
    {
        public SettingsControl()
        {
            InitializeComponent();

            SetText();
        }

        private void SetText()
        {
            GlobalSettingsBox.Text = RubberduckUI.SourceControl_GlobalSettings;
            UserNameLabel.Text = RubberduckUI.SourceControl_UserNameLabel;
            EmailAddressLabel.Text = RubberduckUI.SourceControl_EmailAddressLabel;
            DefaultRepositoryLocationLabel.Text = RubberduckUI.SourceControl_DefaultRepoLocationLabel;
            UpdateGlobalSettingsButton.Text = RubberduckUI.SourceControl_UpdateGlobalSettings;
            CancelGlobalSettingsButton.Text = RubberduckUI.SourceControl_CancelGlobalSettings;

            RepositorySettingsBox.Text = RubberduckUI.SourceControl_RespositorySettings;
            IgnoreFileButton.Text = RubberduckUI.SourceControl_IgnoreFile;
            AttributesFileButton.Text = RubberduckUI.SourceControl_AttributesFile;
        }
    }
}
