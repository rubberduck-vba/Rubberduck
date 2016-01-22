using System;
using System.Diagnostics.CodeAnalysis;
using System.Windows.Forms;
using Rubberduck.Settings;

namespace Rubberduck.UI.SourceControl
{
    [ExcludeFromCodeCoverage]
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
            EditIgnoreFileButton.Text = RubberduckUI.SourceControl_IgnoreFile;
            EditAttributeFileButton.Text = RubberduckUI.SourceControl_AttributesFile;
        }

        string ISourceControlUserSettings.UserName
        {
            get { return UserNameTextBox.Text; }
            set { UserNameTextBox.Text = value; }
        }

        string ISourceControlUserSettings.EmailAddress
        {
            get { return EmailAddressTextBox.Text; }
            set { EmailAddressTextBox.Text = value; }
        }

        string ISourceControlUserSettings.DefaultRepositoryLocation
        {
            get { return DefaultRepositoryLocationTextBox.Text; }
            set { DefaultRepositoryLocationTextBox.Text = value; }
        }

        public event EventHandler<EventArgs> BrowseDefaultRepositoryLocation;
        private void BrowseDefaultRepositoryLocationButton_Click(object sender, EventArgs e)
        {
            var handler = BrowseDefaultRepositoryLocation;
            if (handler != null)
            {
                handler(this, e);
            }
        }

        public event EventHandler<EventArgs> Save;
        private void UpdateGlobalSettingsButton_Click(object sender, EventArgs e)
        {
            var handler = Save;
            if (handler != null)
            {
                handler(this, e);
            }
        }

        public event EventHandler<EventArgs> Cancel;

        private void CancelGlobalSettingsButton_Click(object sender, EventArgs e)
        {
            var handler = Cancel;
            if (handler != null)
            {
                handler(this, e);
            }
        }

        public event EventHandler<EventArgs> EditIgnoreFile;
        private void EditIgnoreFileButton_Click(object sender, EventArgs e)
        {
            var handler = EditIgnoreFile;
            if (handler != null)
            {
                handler(this, e);
            }
        }

        public event EventHandler<EventArgs> EditAttributesFile;
        private void EditAttributeButton_Click(object sender, EventArgs e)
        {
            var handler = EditAttributesFile;
            if (handler != null)
            {
                handler(this, e);
            }
        }
    }
}
