namespace Rubberduck.UI.SourceControl
{
    partial class SourceControlPanel
    {
        /// <summary> 
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

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
            base.Dispose(disposing);
        }

        #region Component Designer generated code

        /// <summary> 
        /// Required method for Designer support - do not modify 
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(SourceControlPanel));
            this.SourceControlToolbar = new System.Windows.Forms.ToolStrip();
            this.RefreshButton = new System.Windows.Forms.ToolStripButton();
            this.OpenWorkingFolderButton = new System.Windows.Forms.ToolStripButton();
            this.toolStripSeparator1 = new System.Windows.Forms.ToolStripSeparator();
            this.StatusMessage = new System.Windows.Forms.ToolStripLabel();
            this.SourceControlTabs = new System.Windows.Forms.TabControl();
            this.ChangesTab = new System.Windows.Forms.TabPage();
            this.BranchesTab = new System.Windows.Forms.TabPage();
            this.UnsyncedCommitsTab = new System.Windows.Forms.TabPage();
            this.SettingsTab = new System.Windows.Forms.TabPage();
            this.InitRepoButton = new System.Windows.Forms.ToolStripButton();
            this.SourceControlToolbar.SuspendLayout();
            this.SourceControlTabs.SuspendLayout();
            this.SuspendLayout();
            // 
            // SourceControlToolbar
            // 
            this.SourceControlToolbar.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.RefreshButton,
            this.OpenWorkingFolderButton,
            this.InitRepoButton,
            this.toolStripSeparator1,
            this.StatusMessage});
            this.SourceControlToolbar.Location = new System.Drawing.Point(0, 0);
            this.SourceControlToolbar.MaximumSize = new System.Drawing.Size(255, 25);
            this.SourceControlToolbar.Name = "SourceControlToolbar";
            this.SourceControlToolbar.Size = new System.Drawing.Size(255, 25);
            this.SourceControlToolbar.TabIndex = 0;
            this.SourceControlToolbar.Text = "toolStrip1";
            // 
            // RefreshButton
            // 
            this.RefreshButton.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.RefreshButton.Image = global::Rubberduck.Properties.Resources.arrow_circle_double;
            this.RefreshButton.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.RefreshButton.Name = "RefreshButton";
            this.RefreshButton.Size = new System.Drawing.Size(23, 22);
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
            this.OpenWorkingFolderButton.Size = new System.Drawing.Size(23, 22);
            this.OpenWorkingFolderButton.ToolTipText = "Open working folder";
            this.OpenWorkingFolderButton.Click += new System.EventHandler(this.OpenWorkingFolderButton_Click);
            // 
            // toolStripSeparator1
            // 
            this.toolStripSeparator1.Name = "toolStripSeparator1";
            this.toolStripSeparator1.Size = new System.Drawing.Size(6, 25);
            // 
            // StatusMessage
            // 
            this.StatusMessage.Enabled = false;
            this.StatusMessage.ForeColor = System.Drawing.SystemColors.ButtonShadow;
            this.StatusMessage.Image = global::Rubberduck.Properties.Resources.icon_github;
            this.StatusMessage.Name = "StatusMessage";
            this.StatusMessage.Size = new System.Drawing.Size(59, 22);
            this.StatusMessage.Text = "Offline";
            // 
            // SourceControlTabs
            // 
            this.SourceControlTabs.Controls.Add(this.ChangesTab);
            this.SourceControlTabs.Controls.Add(this.BranchesTab);
            this.SourceControlTabs.Controls.Add(this.UnsyncedCommitsTab);
            this.SourceControlTabs.Controls.Add(this.SettingsTab);
            this.SourceControlTabs.Dock = System.Windows.Forms.DockStyle.Fill;
            this.SourceControlTabs.Location = new System.Drawing.Point(0, 25);
            this.SourceControlTabs.Name = "SourceControlTabs";
            this.SourceControlTabs.SelectedIndex = 0;
            this.SourceControlTabs.Size = new System.Drawing.Size(255, 449);
            this.SourceControlTabs.TabIndex = 1;
            // 
            // ChangesTab
            // 
            this.ChangesTab.Location = new System.Drawing.Point(4, 22);
            this.ChangesTab.Name = "ChangesTab";
            this.ChangesTab.Padding = new System.Windows.Forms.Padding(3);
            this.ChangesTab.Size = new System.Drawing.Size(247, 423);
            this.ChangesTab.TabIndex = 0;
            this.ChangesTab.Text = "Changes";
            this.ChangesTab.UseVisualStyleBackColor = true;
            // 
            // BranchesTab
            // 
            this.BranchesTab.Location = new System.Drawing.Point(4, 22);
            this.BranchesTab.Name = "BranchesTab";
            this.BranchesTab.Padding = new System.Windows.Forms.Padding(3);
            this.BranchesTab.Size = new System.Drawing.Size(247, 423);
            this.BranchesTab.TabIndex = 1;
            this.BranchesTab.Text = "Branches";
            this.BranchesTab.UseVisualStyleBackColor = true;
            // 
            // UnsyncedCommitsTab
            // 
            this.UnsyncedCommitsTab.Location = new System.Drawing.Point(4, 22);
            this.UnsyncedCommitsTab.Name = "UnsyncedCommitsTab";
            this.UnsyncedCommitsTab.Padding = new System.Windows.Forms.Padding(3);
            this.UnsyncedCommitsTab.Size = new System.Drawing.Size(247, 423);
            this.UnsyncedCommitsTab.TabIndex = 2;
            this.UnsyncedCommitsTab.Text = "Unsynced commits";
            this.UnsyncedCommitsTab.UseVisualStyleBackColor = true;
            // 
            // SettingsTab
            // 
            this.SettingsTab.Location = new System.Drawing.Point(4, 22);
            this.SettingsTab.Name = "SettingsTab";
            this.SettingsTab.Padding = new System.Windows.Forms.Padding(3);
            this.SettingsTab.Size = new System.Drawing.Size(247, 423);
            this.SettingsTab.TabIndex = 3;
            this.SettingsTab.Text = "Settings";
            this.SettingsTab.UseVisualStyleBackColor = true;
            // 
            // InitRepoButton
            // 
            this.InitRepoButton.AccessibleDescription = "Initialize repository from the active project.";
            this.InitRepoButton.AccessibleName = "Initalize Report Button";
            this.InitRepoButton.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.InitRepoButton.Image = ((System.Drawing.Image)(resources.GetObject("InitRepoButton.Image")));
            this.InitRepoButton.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.InitRepoButton.Name = "InitRepoButton";
            this.InitRepoButton.Size = new System.Drawing.Size(23, 22);
            this.InitRepoButton.ToolTipText = "Init New Repo from this Project";
            this.InitRepoButton.Click += new System.EventHandler(this.InitRepoButton_Click);
            // 
            // SourceControlPanel
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.SourceControlTabs);
            this.Controls.Add(this.SourceControlToolbar);
            this.MinimumSize = new System.Drawing.Size(255, 255);
            this.Name = "SourceControlPanel";
            this.Size = new System.Drawing.Size(255, 474);
            this.SourceControlToolbar.ResumeLayout(false);
            this.SourceControlToolbar.PerformLayout();
            this.SourceControlTabs.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.ToolStrip SourceControlToolbar;
        private System.Windows.Forms.ToolStripButton RefreshButton;
        private System.Windows.Forms.TabControl SourceControlTabs;
        private System.Windows.Forms.TabPage ChangesTab;
        private System.Windows.Forms.TabPage BranchesTab;
        private System.Windows.Forms.TabPage UnsyncedCommitsTab;
        private System.Windows.Forms.TabPage SettingsTab;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator1;
        private System.Windows.Forms.ToolStripLabel StatusMessage;
        private System.Windows.Forms.ToolStripButton OpenWorkingFolderButton;
        private System.Windows.Forms.ToolStripButton InitRepoButton;
    }
}
