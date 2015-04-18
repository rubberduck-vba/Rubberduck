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
            this.SourceControlToolbar = new System.Windows.Forms.ToolStrip();
            this.RefreshButton = new System.Windows.Forms.ToolStripButton();
            this.OpenWorkingFolderButton = new System.Windows.Forms.ToolStripButton();
            this.toolStripSeparator1 = new System.Windows.Forms.ToolStripSeparator();
            this.StatusMessage = new System.Windows.Forms.ToolStripLabel();
            this.SourceControlTabs = new System.Windows.Forms.TabControl();
            this.ChangesTab = new System.Windows.Forms.TabPage();
            this.BranchesTab = new System.Windows.Forms.TabPage();
            this.UnsyncedCommitsTab = new System.Windows.Forms.TabPage();
            this.UnsyncedCommitsPanel = new System.Windows.Forms.Panel();
            this.UnsyncedCommitsBranchNameLabel = new System.Windows.Forms.Label();
            this.SyncButton = new System.Windows.Forms.Button();
            this.FetchIncomingCommitsButton = new System.Windows.Forms.Button();
            this.label3 = new System.Windows.Forms.Label();
            this.PushButton = new System.Windows.Forms.Button();
            this.PullButton = new System.Windows.Forms.Button();
            this.OutgoingCommitsBox = new System.Windows.Forms.GroupBox();
            this.OutgoingCommitsGrid = new System.Windows.Forms.DataGridView();
            this.IncomingCommitsBox = new System.Windows.Forms.GroupBox();
            this.IncomingCommitsGrid = new System.Windows.Forms.DataGridView();
            this.SettingsTab = new System.Windows.Forms.TabPage();
            this.SettingsPanel = new System.Windows.Forms.Panel();
            this.RepositorySettingsBox = new System.Windows.Forms.GroupBox();
            this.button2 = new System.Windows.Forms.Button();
            this.button1 = new System.Windows.Forms.Button();
            this.GlobalSettingsBox = new System.Windows.Forms.GroupBox();
            this.CancelGlobalSettingsButton = new System.Windows.Forms.Button();
            this.UpdateGlobalSettingsButton = new System.Windows.Forms.Button();
            this.BrowseDefaultRepositoryLocationButton = new System.Windows.Forms.Button();
            this.DefaultRepositoryLocation = new System.Windows.Forms.TextBox();
            this.label7 = new System.Windows.Forms.Label();
            this.EmailAddress = new System.Windows.Forms.TextBox();
            this.label6 = new System.Windows.Forms.Label();
            this.UserName = new System.Windows.Forms.TextBox();
            this.label5 = new System.Windows.Forms.Label();
            this.SourceControlToolbar.SuspendLayout();
            this.SourceControlTabs.SuspendLayout();
            this.UnsyncedCommitsTab.SuspendLayout();
            this.UnsyncedCommitsPanel.SuspendLayout();
            this.OutgoingCommitsBox.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.OutgoingCommitsGrid)).BeginInit();
            this.IncomingCommitsBox.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.IncomingCommitsGrid)).BeginInit();
            this.SettingsTab.SuspendLayout();
            this.SettingsPanel.SuspendLayout();
            this.RepositorySettingsBox.SuspendLayout();
            this.GlobalSettingsBox.SuspendLayout();
            this.SuspendLayout();
            // 
            // SourceControlToolbar
            // 
            this.SourceControlToolbar.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.RefreshButton,
            this.OpenWorkingFolderButton,
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
            this.UnsyncedCommitsTab.Controls.Add(this.UnsyncedCommitsPanel);
            this.UnsyncedCommitsTab.Location = new System.Drawing.Point(4, 22);
            this.UnsyncedCommitsTab.Name = "UnsyncedCommitsTab";
            this.UnsyncedCommitsTab.Padding = new System.Windows.Forms.Padding(3);
            this.UnsyncedCommitsTab.Size = new System.Drawing.Size(247, 423);
            this.UnsyncedCommitsTab.TabIndex = 2;
            this.UnsyncedCommitsTab.Text = "Unsynced commits";
            this.UnsyncedCommitsTab.UseVisualStyleBackColor = true;
            // 
            // UnsyncedCommitsPanel
            // 
            this.UnsyncedCommitsPanel.AutoScroll = true;
            this.UnsyncedCommitsPanel.Controls.Add(this.UnsyncedCommitsBranchNameLabel);
            this.UnsyncedCommitsPanel.Controls.Add(this.SyncButton);
            this.UnsyncedCommitsPanel.Controls.Add(this.FetchIncomingCommitsButton);
            this.UnsyncedCommitsPanel.Controls.Add(this.label3);
            this.UnsyncedCommitsPanel.Controls.Add(this.PushButton);
            this.UnsyncedCommitsPanel.Controls.Add(this.PullButton);
            this.UnsyncedCommitsPanel.Controls.Add(this.OutgoingCommitsBox);
            this.UnsyncedCommitsPanel.Controls.Add(this.IncomingCommitsBox);
            this.UnsyncedCommitsPanel.Dock = System.Windows.Forms.DockStyle.Fill;
            this.UnsyncedCommitsPanel.Location = new System.Drawing.Point(3, 3);
            this.UnsyncedCommitsPanel.Name = "UnsyncedCommitsPanel";
            this.UnsyncedCommitsPanel.Padding = new System.Windows.Forms.Padding(3);
            this.UnsyncedCommitsPanel.Size = new System.Drawing.Size(241, 417);
            this.UnsyncedCommitsPanel.TabIndex = 0;
            // 
            // UnsyncedCommitsBranchNameLabel
            // 
            this.UnsyncedCommitsBranchNameLabel.AutoSize = true;
            this.UnsyncedCommitsBranchNameLabel.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.UnsyncedCommitsBranchNameLabel.Location = new System.Drawing.Point(56, 14);
            this.UnsyncedCommitsBranchNameLabel.Name = "UnsyncedCommitsBranchNameLabel";
            this.UnsyncedCommitsBranchNameLabel.Size = new System.Drawing.Size(45, 13);
            this.UnsyncedCommitsBranchNameLabel.TabIndex = 19;
            this.UnsyncedCommitsBranchNameLabel.Text = "Master";
            // 
            // SyncButton
            // 
            this.SyncButton.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.SyncButton.Location = new System.Drawing.Point(9, 68);
            this.SyncButton.Name = "SyncButton";
            this.SyncButton.Size = new System.Drawing.Size(201, 23);
            this.SyncButton.TabIndex = 11;
            this.SyncButton.Text = "Sync";
            this.SyncButton.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText;
            this.SyncButton.UseVisualStyleBackColor = true;
            // 
            // FetchIncomingCommitsButton
            // 
            this.FetchIncomingCommitsButton.Image = global::Rubberduck.Properties.Resources.arrow_step;
            this.FetchIncomingCommitsButton.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.FetchIncomingCommitsButton.Location = new System.Drawing.Point(9, 39);
            this.FetchIncomingCommitsButton.Name = "FetchIncomingCommitsButton";
            this.FetchIncomingCommitsButton.Size = new System.Drawing.Size(63, 23);
            this.FetchIncomingCommitsButton.TabIndex = 13;
            this.FetchIncomingCommitsButton.Text = "Fetch";
            this.FetchIncomingCommitsButton.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText;
            this.FetchIncomingCommitsButton.UseVisualStyleBackColor = true;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(6, 14);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(44, 13);
            this.label3.TabIndex = 9;
            this.label3.Text = "Branch:";
            // 
            // PushButton
            // 
            this.PushButton.Image = global::Rubberduck.Properties.Resources.drive_upload;
            this.PushButton.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.PushButton.Location = new System.Drawing.Point(147, 39);
            this.PushButton.Name = "PushButton";
            this.PushButton.Size = new System.Drawing.Size(63, 23);
            this.PushButton.TabIndex = 14;
            this.PushButton.Text = "Push";
            this.PushButton.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText;
            this.PushButton.UseVisualStyleBackColor = true;
            // 
            // PullButton
            // 
            this.PullButton.Image = global::Rubberduck.Properties.Resources.drive_download;
            this.PullButton.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.PullButton.Location = new System.Drawing.Point(78, 39);
            this.PullButton.Name = "PullButton";
            this.PullButton.Size = new System.Drawing.Size(63, 23);
            this.PullButton.TabIndex = 12;
            this.PullButton.Text = "Pull";
            this.PullButton.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText;
            this.PullButton.UseVisualStyleBackColor = true;
            // 
            // OutgoingCommitsBox
            // 
            this.OutgoingCommitsBox.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.OutgoingCommitsBox.Controls.Add(this.OutgoingCommitsGrid);
            this.OutgoingCommitsBox.Location = new System.Drawing.Point(9, 269);
            this.OutgoingCommitsBox.Name = "OutgoingCommitsBox";
            this.OutgoingCommitsBox.Padding = new System.Windows.Forms.Padding(6);
            this.OutgoingCommitsBox.Size = new System.Drawing.Size(50, 162);
            this.OutgoingCommitsBox.TabIndex = 16;
            this.OutgoingCommitsBox.TabStop = false;
            this.OutgoingCommitsBox.Text = "Outgoing commits";
            // 
            // OutgoingCommitsGrid
            // 
            this.OutgoingCommitsGrid.AllowUserToAddRows = false;
            this.OutgoingCommitsGrid.AllowUserToDeleteRows = false;
            this.OutgoingCommitsGrid.AllowUserToResizeColumns = false;
            this.OutgoingCommitsGrid.AllowUserToResizeRows = false;
            this.OutgoingCommitsGrid.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.OutgoingCommitsGrid.BackgroundColor = System.Drawing.Color.White;
            this.OutgoingCommitsGrid.CellBorderStyle = System.Windows.Forms.DataGridViewCellBorderStyle.None;
            this.OutgoingCommitsGrid.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.OutgoingCommitsGrid.EditMode = System.Windows.Forms.DataGridViewEditMode.EditProgrammatically;
            this.OutgoingCommitsGrid.GridColor = System.Drawing.Color.White;
            this.OutgoingCommitsGrid.Location = new System.Drawing.Point(10, 22);
            this.OutgoingCommitsGrid.Name = "OutgoingCommitsGrid";
            this.OutgoingCommitsGrid.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.OutgoingCommitsGrid.Size = new System.Drawing.Size(31, 131);
            this.OutgoingCommitsGrid.TabIndex = 0;
            // 
            // IncomingCommitsBox
            // 
            this.IncomingCommitsBox.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.IncomingCommitsBox.Controls.Add(this.IncomingCommitsGrid);
            this.IncomingCommitsBox.Location = new System.Drawing.Point(9, 101);
            this.IncomingCommitsBox.Name = "IncomingCommitsBox";
            this.IncomingCommitsBox.Padding = new System.Windows.Forms.Padding(6);
            this.IncomingCommitsBox.Size = new System.Drawing.Size(50, 162);
            this.IncomingCommitsBox.TabIndex = 15;
            this.IncomingCommitsBox.TabStop = false;
            this.IncomingCommitsBox.Text = "Incoming commits";
            // 
            // IncomingCommitsGrid
            // 
            this.IncomingCommitsGrid.AllowUserToAddRows = false;
            this.IncomingCommitsGrid.AllowUserToDeleteRows = false;
            this.IncomingCommitsGrid.AllowUserToResizeColumns = false;
            this.IncomingCommitsGrid.AllowUserToResizeRows = false;
            this.IncomingCommitsGrid.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.IncomingCommitsGrid.BackgroundColor = System.Drawing.Color.White;
            this.IncomingCommitsGrid.CellBorderStyle = System.Windows.Forms.DataGridViewCellBorderStyle.None;
            this.IncomingCommitsGrid.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.IncomingCommitsGrid.EditMode = System.Windows.Forms.DataGridViewEditMode.EditProgrammatically;
            this.IncomingCommitsGrid.GridColor = System.Drawing.Color.White;
            this.IncomingCommitsGrid.Location = new System.Drawing.Point(10, 22);
            this.IncomingCommitsGrid.Name = "IncomingCommitsGrid";
            this.IncomingCommitsGrid.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.IncomingCommitsGrid.Size = new System.Drawing.Size(31, 131);
            this.IncomingCommitsGrid.TabIndex = 0;
            // 
            // SettingsTab
            // 
            this.SettingsTab.Controls.Add(this.SettingsPanel);
            this.SettingsTab.Location = new System.Drawing.Point(4, 22);
            this.SettingsTab.Name = "SettingsTab";
            this.SettingsTab.Padding = new System.Windows.Forms.Padding(3);
            this.SettingsTab.Size = new System.Drawing.Size(247, 423);
            this.SettingsTab.TabIndex = 3;
            this.SettingsTab.Text = "Settings";
            this.SettingsTab.UseVisualStyleBackColor = true;
            // 
            // SettingsPanel
            // 
            this.SettingsPanel.AutoScroll = true;
            this.SettingsPanel.Controls.Add(this.RepositorySettingsBox);
            this.SettingsPanel.Controls.Add(this.GlobalSettingsBox);
            this.SettingsPanel.Dock = System.Windows.Forms.DockStyle.Fill;
            this.SettingsPanel.Location = new System.Drawing.Point(3, 3);
            this.SettingsPanel.Name = "SettingsPanel";
            this.SettingsPanel.Padding = new System.Windows.Forms.Padding(3);
            this.SettingsPanel.Size = new System.Drawing.Size(241, 417);
            this.SettingsPanel.TabIndex = 0;
            // 
            // RepositorySettingsBox
            // 
            this.RepositorySettingsBox.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.RepositorySettingsBox.Controls.Add(this.button2);
            this.RepositorySettingsBox.Controls.Add(this.button1);
            this.RepositorySettingsBox.Location = new System.Drawing.Point(6, 212);
            this.RepositorySettingsBox.Name = "RepositorySettingsBox";
            this.RepositorySettingsBox.Padding = new System.Windows.Forms.Padding(6);
            this.RepositorySettingsBox.Size = new System.Drawing.Size(228, 71);
            this.RepositorySettingsBox.TabIndex = 3;
            this.RepositorySettingsBox.TabStop = false;
            this.RepositorySettingsBox.Text = "Repository Settings";
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(105, 31);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(92, 23);
            this.button2.TabIndex = 1;
            this.button2.Text = "Attributes File";
            this.button2.UseVisualStyleBackColor = true;
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(7, 31);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(92, 23);
            this.button1.TabIndex = 0;
            this.button1.Text = "Ignore File";
            this.button1.UseVisualStyleBackColor = true;
            // 
            // GlobalSettingsBox
            // 
            this.GlobalSettingsBox.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.GlobalSettingsBox.Controls.Add(this.CancelGlobalSettingsButton);
            this.GlobalSettingsBox.Controls.Add(this.UpdateGlobalSettingsButton);
            this.GlobalSettingsBox.Controls.Add(this.BrowseDefaultRepositoryLocationButton);
            this.GlobalSettingsBox.Controls.Add(this.DefaultRepositoryLocation);
            this.GlobalSettingsBox.Controls.Add(this.label7);
            this.GlobalSettingsBox.Controls.Add(this.EmailAddress);
            this.GlobalSettingsBox.Controls.Add(this.label6);
            this.GlobalSettingsBox.Controls.Add(this.UserName);
            this.GlobalSettingsBox.Controls.Add(this.label5);
            this.GlobalSettingsBox.Location = new System.Drawing.Point(6, 6);
            this.GlobalSettingsBox.Name = "GlobalSettingsBox";
            this.GlobalSettingsBox.Padding = new System.Windows.Forms.Padding(6);
            this.GlobalSettingsBox.Size = new System.Drawing.Size(228, 199);
            this.GlobalSettingsBox.TabIndex = 2;
            this.GlobalSettingsBox.TabStop = false;
            this.GlobalSettingsBox.Text = "Global Settings";
            // 
            // CancelGlobalSettingsButton
            // 
            this.CancelGlobalSettingsButton.Location = new System.Drawing.Point(105, 168);
            this.CancelGlobalSettingsButton.Name = "CancelGlobalSettingsButton";
            this.CancelGlobalSettingsButton.Size = new System.Drawing.Size(92, 23);
            this.CancelGlobalSettingsButton.TabIndex = 8;
            this.CancelGlobalSettingsButton.Text = "Cancel";
            this.CancelGlobalSettingsButton.UseVisualStyleBackColor = true;
            // 
            // UpdateGlobalSettingsButton
            // 
            this.UpdateGlobalSettingsButton.Location = new System.Drawing.Point(7, 168);
            this.UpdateGlobalSettingsButton.Name = "UpdateGlobalSettingsButton";
            this.UpdateGlobalSettingsButton.Size = new System.Drawing.Size(92, 23);
            this.UpdateGlobalSettingsButton.TabIndex = 7;
            this.UpdateGlobalSettingsButton.Text = "Update";
            this.UpdateGlobalSettingsButton.UseVisualStyleBackColor = true;
            // 
            // BrowseDefaultRepositoryLocationButton
            // 
            this.BrowseDefaultRepositoryLocationButton.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.BrowseDefaultRepositoryLocationButton.Location = new System.Drawing.Point(184, 120);
            this.BrowseDefaultRepositoryLocationButton.Name = "BrowseDefaultRepositoryLocationButton";
            this.BrowseDefaultRepositoryLocationButton.Size = new System.Drawing.Size(33, 20);
            this.BrowseDefaultRepositoryLocationButton.TabIndex = 6;
            this.BrowseDefaultRepositoryLocationButton.Text = "...";
            this.BrowseDefaultRepositoryLocationButton.UseVisualStyleBackColor = true;
            // 
            // DefaultRepositoryLocation
            // 
            this.DefaultRepositoryLocation.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.DefaultRepositoryLocation.Location = new System.Drawing.Point(12, 120);
            this.DefaultRepositoryLocation.Name = "DefaultRepositoryLocation";
            this.DefaultRepositoryLocation.Size = new System.Drawing.Size(173, 20);
            this.DefaultRepositoryLocation.TabIndex = 5;
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(9, 103);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(138, 13);
            this.label7.TabIndex = 4;
            this.label7.Text = "Default Repository Location";
            // 
            // EmailAddress
            // 
            this.EmailAddress.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.EmailAddress.Location = new System.Drawing.Point(12, 80);
            this.EmailAddress.Name = "EmailAddress";
            this.EmailAddress.Size = new System.Drawing.Size(206, 20);
            this.EmailAddress.TabIndex = 3;
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(9, 63);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(73, 13);
            this.label6.TabIndex = 2;
            this.label6.Text = "Email Address";
            // 
            // UserName
            // 
            this.UserName.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.UserName.Location = new System.Drawing.Point(13, 40);
            this.UserName.Name = "UserName";
            this.UserName.Size = new System.Drawing.Size(206, 20);
            this.UserName.TabIndex = 1;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(10, 23);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(60, 13);
            this.label5.TabIndex = 0;
            this.label5.Text = "User Name";
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
            this.UnsyncedCommitsTab.ResumeLayout(false);
            this.UnsyncedCommitsPanel.ResumeLayout(false);
            this.UnsyncedCommitsPanel.PerformLayout();
            this.OutgoingCommitsBox.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.OutgoingCommitsGrid)).EndInit();
            this.IncomingCommitsBox.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.IncomingCommitsGrid)).EndInit();
            this.SettingsTab.ResumeLayout(false);
            this.SettingsPanel.ResumeLayout(false);
            this.RepositorySettingsBox.ResumeLayout(false);
            this.GlobalSettingsBox.ResumeLayout(false);
            this.GlobalSettingsBox.PerformLayout();
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
        private System.Windows.Forms.Panel UnsyncedCommitsPanel;
        private System.Windows.Forms.Button FetchIncomingCommitsButton;
        private System.Windows.Forms.Button SyncButton;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Button PushButton;
        private System.Windows.Forms.Button PullButton;
        private System.Windows.Forms.GroupBox OutgoingCommitsBox;
        private System.Windows.Forms.DataGridView OutgoingCommitsGrid;
        private System.Windows.Forms.GroupBox IncomingCommitsBox;
        private System.Windows.Forms.DataGridView IncomingCommitsGrid;
        private System.Windows.Forms.Panel SettingsPanel;
        private System.Windows.Forms.GroupBox RepositorySettingsBox;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.GroupBox GlobalSettingsBox;
        private System.Windows.Forms.Button CancelGlobalSettingsButton;
        private System.Windows.Forms.Button UpdateGlobalSettingsButton;
        private System.Windows.Forms.Button BrowseDefaultRepositoryLocationButton;
        private System.Windows.Forms.TextBox DefaultRepositoryLocation;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.TextBox EmailAddress;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.TextBox UserName;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.ToolStripButton OpenWorkingFolderButton;
        private System.Windows.Forms.Label UnsyncedCommitsBranchNameLabel;
    }
}
