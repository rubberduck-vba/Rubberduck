using System;
using System.Windows.Forms;
using Rubberduck.Config;

namespace Rubberduck.UI.SourceControl
{
    public partial class SettingsControl : UserControl, ISettingsView
    {
        public SettingsControl()
        {
            InitializeComponent();
        }

        string ISourceControlUserSettings.UserName
        {
            get { return this.UserName.Text; }
            set { this.UserName.Text = value; }
        }

        string ISourceControlUserSettings.EmailAddress
        {
            get { return this.EmailAddress.Text; }
            set { this.EmailAddress.Text = value; }
        }

        string ISourceControlUserSettings.DefaultRepositoryLocation
        {
            get { return this.DefaultRepositoryLocation.Text; }
            set { this.DefaultRepositoryLocation.Text = value; }
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
