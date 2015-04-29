using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics.CodeAnalysis;
using System.Drawing;
using System.Windows.Forms;
using Rubberduck.SourceControl;

namespace Rubberduck.UI.SourceControl
{
    [SuppressMessage("ReSharper", "ArrangeThisQualifier")]
    public partial class SourceControlPanel : UserControl, ISourceControlView
    {
        public SourceControlPanel()
        {
            InitializeComponent();
        }

        public SourceControlPanel(IBranchesView branchesView, IChangesView changesView, IUnSyncedCommitsView commitsView, ISettingsView settingsView)
            :this()
        {
            this.BranchesTab.Controls.Add((Control)branchesView);
            this.ChangesTab.Controls.Add((Control)changesView);
            this.UnsyncedCommitsTab.Controls.Add((Control)commitsView);
            this.SettingsTab.Controls.Add((Control)settingsView);
        }

        public string ClassId
        {
            get { return "19A32FC9-4902-4385-9FE7-829D4F9C441D"; }
        }

        public string Caption
        {
            get { return "Source Control"; }
        }

        public string Status 
        {
            get { return this.StatusMessage.Text; }
            set { this.StatusMessage.Text = value; }
        }

        public event EventHandler<EventArgs> RefreshData;
        private void RefreshButton_Click(object sender, EventArgs e)
        {
            RaiseGenericEvent(RefreshData, e);
        }

        private void RaiseGenericEvent(EventHandler<EventArgs> handler, EventArgs e)
        {
            if (handler != null)
            {
                handler(this, e);
            }
        }
    }
}
