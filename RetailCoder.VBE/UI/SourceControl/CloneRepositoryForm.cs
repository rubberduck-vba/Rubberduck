using System;
using System.Diagnostics.CodeAnalysis;
using System.Windows.Forms;

namespace Rubberduck.UI.SourceControl
{
    [ExcludeFromCodeCoverage]
    public partial class CloneRepositoryForm : Form, ICloneRepositoryView
    {
        public CloneRepositoryForm()
        {
            InitializeComponent();

            RemotePathTextBox.TextChanged += RemotePath_TextChanged;
            BrowseLocalDirectoryLocationButton.Click += OnBrowseLocalDirectoryLocation;
            OkButton.Click += OkButton_Click;
            CancelButton.Click += CancelButton_Click;
        }

        public CloneRepositoryForm(IFolderBrowserFactory folderBrowserFactory)
            : this()
        {
            FolderBrowserFactory = folderBrowserFactory;
        }

        public IFolderBrowserFactory FolderBrowserFactory { get; set; }

        public string RemotePath
        {
            get { return RemotePathTextBox.Text; }
            set { RemotePathTextBox.Text = value; }
        }

        public string LocalDirectory
        {
            get { return LocalDirectoryTextBox.Text; }
            set { LocalDirectoryTextBox.Text = value; }
        }

        public bool IsValidRemotePath
        {
            get { return InvalidRemotePathValidationIcon.Visible; }
            set
            {
                OkButton.Enabled = value;
                InvalidRemotePathValidationIcon.Visible = !value;
            }
        }

        public event EventHandler<CloneRepositoryEventArgs> Confirm;
        private void OkButton_Click(object sender, EventArgs e)
        {
            var handler = Confirm;
            if (handler != null)
            {
                handler(this, new CloneRepositoryEventArgs(RemotePath, LocalDirectory));
            }
        }

        public event EventHandler<EventArgs> Cancel;
        private void CancelButton_Click(object sender, EventArgs e)
        {
            var handler = Cancel;
            if (handler != null)
            {
                handler(this, e);
            }
        }

        public event EventHandler<EventArgs> RemotePathChanged;
        private void RemotePath_TextChanged(object sender, EventArgs e)
        {
            var handler = RemotePathChanged;
            if (handler != null)
            {
                handler(this, e);
            }

        }

        private void OnBrowseLocalDirectoryLocation(object sender, EventArgs e)
        {
            using (var folderPicker = FolderBrowserFactory.CreateFolderBrowser("Local Directory"))
            {
                if (folderPicker.ShowDialog() == DialogResult.OK)
                {
                    LocalDirectory = folderPicker.SelectedPath;
                }
            }
        }
    }
}
