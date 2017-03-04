using System.ComponentModel;
using System.Windows.Forms;

namespace Rubberduck.UI.SourceControl
{
    partial class SourceControlPanel
    {
        /// <summary> 
        /// Required designer variable.
        /// </summary>
        private IContainer components = null;

        /// <summary> 
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
        }

        #region Component Designer generated code

        /// <summary> 
        /// Required method for Designer support - do not modify 
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            /*System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(SourceControlPanel));
            this.SourceControlToolbar = new System.Windows.Forms.ToolStrip();
            this.RefreshButton = new System.Windows.Forms.ToolStripButton();
            this.OpenWorkingFolderButton = new System.Windows.Forms.ToolStripButton();
            this.InitRepoButton = new System.Windows.Forms.ToolStripButton();
            this.CloneRepoButton = new System.Windows.Forms.ToolStripButton();
            this.toolStripSeparator1 = new System.Windows.Forms.ToolStripSeparator();
            this.StatusMessage = new System.Windows.Forms.ToolStripLabel();
            this.SourceControlTabs = new System.Windows.Forms.TabControl();
            this.ChangesTab = new System.Windows.Forms.TabPage();
            this.BranchesTab = new System.Windows.Forms.TabPage();
            this.UnsyncedCommitsTab = new System.Windows.Forms.TabPage();
            this.SettingsTab = new System.Windows.Forms.TabPage();
            this.MainContainer = new System.Windows.Forms.SplitContainer();
            this.SourceControlToolbar.SuspendLayout();
            this.SourceControlTabs.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.MainContainer)).BeginInit();
            this.MainContainer.Panel2.SuspendLayout();
            this.MainContainer.SuspendLayout();
            this.SuspendLayout();
            // 
            // SourceControlToolbar
            // 
            this.SourceControlToolbar.ImageScalingSize = new System.Drawing.Size(20, 20);
            this.SourceControlToolbar.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.RefreshButton,
            this.OpenWorkingFolderButton,
            this.InitRepoButton,
            this.CloneRepoButton,
            this.toolStripSeparator1,
            this.StatusMessage});
            this.SourceControlToolbar.Location = new System.Drawing.Point(0, 0);
            this.SourceControlToolbar.MaximumSize = new System.Drawing.Size(340, 31);
            this.SourceControlToolbar.Name = "SourceControlToolbar";
            this.SourceControlToolbar.Size = new System.Drawing.Size(340, 27);
            this.SourceControlToolbar.TabIndex = 0;
            this.SourceControlToolbar.Text = "toolStrip1";
            // 
            // RefreshButton
            // 
            this.RefreshButton.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.RefreshButton.Image = global::Rubberduck.Properties.Resources.arrow_circle_double;
            this.RefreshButton.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.RefreshButton.Name = "RefreshButton";
            this.RefreshButton.Size = new System.Drawing.Size(24, 24);
            this.RefreshButton.Text = "Refresh";
            this.RefreshButton.ToolTipText = "Refreshes pending changes";
            this.RefreshButton.Click += new System.EventHandler(this.RefreshButton_Click);
            // 
            // OpenWorkingFolderButton
            // 
            this.OpenWorkingFolderButton.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.OpenWorkingFolderButton.Image = global::Rubberduck.Properties.Resources.folder_horizontal_open;
            this.OpenWorkingFolderButton.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.OpenWorkingFolderButton.Name = "OpenWorkingFolderButton";
            this.OpenWorkingFolderButton.Size = new System.Drawing.Size(24, 24);
            this.OpenWorkingFolderButton.ToolTipText = "Open working folder";
            this.OpenWorkingFolderButton.Click += new System.EventHandler(this.OpenWorkingFolderButton_Click);
            // 
            // InitRepoButton
            // 
            this.InitRepoButton.AccessibleDescription = "Initialize repository from the active project.";
            this.InitRepoButton.AccessibleName = "Initalize Report Button";
            this.InitRepoButton.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.InitRepoButton.Image = ((System.Drawing.Image)(resources.GetObject("InitRepoButton.Image")));
            this.InitRepoButton.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.InitRepoButton.Name = "InitRepoButton";
            this.InitRepoButton.Size = new System.Drawing.Size(24, 24);
            this.InitRepoButton.ToolTipText = "Init New Repo from this Project";
            this.InitRepoButton.Click += new System.EventHandler(this.InitRepoButton_Click);
            // 
            // CloneRepoButton
            // 
            this.CloneRepoButton.AccessibleRole = System.Windows.Forms.AccessibleRole.None;
            this.CloneRepoButton.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.CloneRepoButton.Image = global::Rubberduck.Properties.Resources.drive_download;
            this.CloneRepoButton.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.CloneRepoButton.Name = "CloneRepoButton";
            this.CloneRepoButton.Size = new System.Drawing.Size(24, 24);
            this.CloneRepoButton.Text = "Clone repo";
            this.CloneRepoButton.Click += new System.EventHandler(this.CloneRepoButton_Click);
            // 
            // toolStripSeparator1
            // 
            this.toolStripSeparator1.Name = "toolStripSeparator1";
            this.toolStripSeparator1.Size = new System.Drawing.Size(6, 27);
            // 
            // StatusMessage
            // 
            this.StatusMessage.Enabled = false;
            this.StatusMessage.ForeColor = System.Drawing.SystemColors.ButtonShadow;
            this.StatusMessage.Image = global::Rubberduck.Properties.Resources.icon_github;
            this.StatusMessage.Name = "StatusMessage";
            this.StatusMessage.Size = new System.Drawing.Size(74, 24);
            this.StatusMessage.Text = "Offline";
            // 
            // SourceControlTabs
            // 
            this.SourceControlTabs.Controls.Add(this.ChangesTab);
            this.SourceControlTabs.Controls.Add(this.BranchesTab);
            this.SourceControlTabs.Controls.Add(this.UnsyncedCommitsTab);
            this.SourceControlTabs.Controls.Add(this.SettingsTab);
            this.SourceControlTabs.Dock = System.Windows.Forms.DockStyle.Fill;
            this.SourceControlTabs.Location = new System.Drawing.Point(0, 0);
            this.SourceControlTabs.Margin = new System.Windows.Forms.Padding(4);
            this.SourceControlTabs.Name = "SourceControlTabs";
            this.SourceControlTabs.SelectedIndex = 0;
            this.SourceControlTabs.Size = new System.Drawing.Size(511, 405);
            this.SourceControlTabs.TabIndex = 1;
            // 
            // ChangesTab
            // 
            this.ChangesTab.BackColor = System.Drawing.Color.Transparent;
            this.ChangesTab.Location = new System.Drawing.Point(4, 25);
            this.ChangesTab.Margin = new System.Windows.Forms.Padding(4);
            this.ChangesTab.Name = "ChangesTab";
            this.ChangesTab.Padding = new System.Windows.Forms.Padding(4);
            this.ChangesTab.Size = new System.Drawing.Size(503, 376);
            this.ChangesTab.TabIndex = 0;
            this.ChangesTab.Text = "Changes";
            // 
            // BranchesTab
            // 
            this.BranchesTab.Location = new System.Drawing.Point(4, 25);
            this.BranchesTab.Margin = new System.Windows.Forms.Padding(4);
            this.BranchesTab.Name = "BranchesTab";
            this.BranchesTab.Padding = new System.Windows.Forms.Padding(4);
            this.BranchesTab.Size = new System.Drawing.Size(503, 376);
            this.BranchesTab.TabIndex = 1;
            this.BranchesTab.Text = "Branches";
            this.BranchesTab.UseVisualStyleBackColor = true;
            // 
            // UnsyncedCommitsTab
            // 
            this.UnsyncedCommitsTab.Location = new System.Drawing.Point(4, 25);
            this.UnsyncedCommitsTab.Margin = new System.Windows.Forms.Padding(4);
            this.UnsyncedCommitsTab.Name = "UnsyncedCommitsTab";
            this.UnsyncedCommitsTab.Padding = new System.Windows.Forms.Padding(4);
            this.UnsyncedCommitsTab.Size = new System.Drawing.Size(503, 376);
            this.UnsyncedCommitsTab.TabIndex = 2;
            this.UnsyncedCommitsTab.Text = "Unsynced commits";
            this.UnsyncedCommitsTab.UseVisualStyleBackColor = true;
            // 
            // SettingsTab
            // 
            this.SettingsTab.Location = new System.Drawing.Point(4, 25);
            this.SettingsTab.Margin = new System.Windows.Forms.Padding(4);
            this.SettingsTab.Name = "SettingsTab";
            this.SettingsTab.Padding = new System.Windows.Forms.Padding(4);
            this.SettingsTab.Size = new System.Drawing.Size(503, 376);
            this.SettingsTab.TabIndex = 3;
            this.SettingsTab.Text = "Settings";
            this.SettingsTab.UseVisualStyleBackColor = true;
            // 
            // MainContainer
            // 
            this.MainContainer.Dock = System.Windows.Forms.DockStyle.Fill;
            this.MainContainer.IsSplitterFixed = true;
            this.MainContainer.Location = new System.Drawing.Point(0, 27);
            this.MainContainer.Margin = new System.Windows.Forms.Padding(4);
            this.MainContainer.Name = "MainContainer";
            this.MainContainer.Orientation = System.Windows.Forms.Orientation.Horizontal;
            // 
            // MainContainer.Panel2
            // 
            this.MainContainer.Panel2.Controls.Add(this.SourceControlTabs);
            this.MainContainer.Size = new System.Drawing.Size(511, 556);
            this.MainContainer.SplitterDistance = 146;
            this.MainContainer.SplitterWidth = 5;
            this.MainContainer.TabIndex = 2;
            // 
            // SourceControlPanel
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.MainContainer);
            this.Controls.Add(this.SourceControlToolbar);
            this.Margin = new System.Windows.Forms.Padding(4);
            this.MinimumSize = new System.Drawing.Size(340, 314);
            this.Name = "SourceControlPanel";
            this.Size = new System.Drawing.Size(511, 583);
            this.SourceControlToolbar.ResumeLayout(false);
            this.SourceControlToolbar.PerformLayout();
            this.SourceControlTabs.ResumeLayout(false);
            this.MainContainer.Panel2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.MainContainer)).EndInit();
            this.MainContainer.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();*/

            this.ElementHost = new System.Windows.Forms.Integration.ElementHost();
            this.SourceControlPanelControl = new Rubberduck.UI.SourceControl.SourceControlView();
            this.SuspendLayout();
            // 
            // elementHost1
            // 
            this.ElementHost.Dock = System.Windows.Forms.DockStyle.Fill;
            this.ElementHost.Location = new System.Drawing.Point(0, 0);
            this.ElementHost.Name = "elementHost1";
            this.ElementHost.Size = new System.Drawing.Size(150, 150);
            this.ElementHost.TabIndex = 0;
            this.ElementHost.Text = "elementHost1";
            this.ElementHost.Child = this.SourceControlPanelControl;
            // 
            // SourceControlWindow
            // 
            this.Controls.Add(this.ElementHost);
            this.Name = "SourceControlWindow";
            this.ResumeLayout(false);
        }

        #endregion

        /*private ToolStrip SourceControlToolbar;
        private ToolStripButton RefreshButton;
        private TabControl SourceControlTabs;
        private TabPage ChangesTab;
        private TabPage BranchesTab;
        private TabPage UnsyncedCommitsTab;
        private TabPage SettingsTab;
        private ToolStripSeparator toolStripSeparator1;
        private ToolStripLabel StatusMessage;
        private ToolStripButton OpenWorkingFolderButton;
        private ToolStripButton InitRepoButton;
        private SplitContainer MainContainer;
        private ToolStripButton CloneRepoButton;*/

        private System.Windows.Forms.Integration.ElementHost ElementHost;
        private SourceControlView SourceControlPanelControl;
    }
}
