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
            this.toolStrip1 = new System.Windows.Forms.ToolStrip();
            this.tabControl1 = new System.Windows.Forms.TabControl();
            this.tabPage1 = new System.Windows.Forms.TabPage();
            this.tabPage2 = new System.Windows.Forms.TabPage();
            this.tabPage3 = new System.Windows.Forms.TabPage();
            this.tabPage4 = new System.Windows.Forms.TabPage();
            this.label1 = new System.Windows.Forms.Label();
            this.CurrentBranch = new System.Windows.Forms.ComboBox();
            this.CommitMessage = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.CommitAction = new System.Windows.Forms.ComboBox();
            this.CommitButton = new System.Windows.Forms.Button();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.IncludedChangesGrid = new System.Windows.Forms.DataGridView();
            this.ExcludedChangesGrid = new System.Windows.Forms.DataGridView();
            this.comboBox1 = new System.Windows.Forms.ComboBox();
            this.label3 = new System.Windows.Forms.Label();
            this.groupBox3 = new System.Windows.Forms.GroupBox();
            this.IncomingCommitsGrid = new System.Windows.Forms.DataGridView();
            this.groupBox4 = new System.Windows.Forms.GroupBox();
            this.OutgoingCommitsGrid = new System.Windows.Forms.DataGridView();
            this.comboBox2 = new System.Windows.Forms.ComboBox();
            this.label4 = new System.Windows.Forms.Label();
            this.groupBox5 = new System.Windows.Forms.GroupBox();
            this.label5 = new System.Windows.Forms.Label();
            this.UserName = new System.Windows.Forms.TextBox();
            this.EmailAddress = new System.Windows.Forms.TextBox();
            this.label6 = new System.Windows.Forms.Label();
            this.DefaultRepositoryLocation = new System.Windows.Forms.TextBox();
            this.label7 = new System.Windows.Forms.Label();
            this.BrowseDefaultRepositoryLocationButton = new System.Windows.Forms.Button();
            this.UpdateGlobalSettingsButton = new System.Windows.Forms.Button();
            this.CancelGlobalSettingsButton = new System.Windows.Forms.Button();
            this.groupBox6 = new System.Windows.Forms.GroupBox();
            this.button1 = new System.Windows.Forms.Button();
            this.button2 = new System.Windows.Forms.Button();
            this.groupBox7 = new System.Windows.Forms.GroupBox();
            this.UnpublishedBranchesGrid = new System.Windows.Forms.DataGridView();
            this.groupBox8 = new System.Windows.Forms.GroupBox();
            this.PublishedBranchesGrid = new System.Windows.Forms.DataGridView();
            this.MergeButton = new System.Windows.Forms.Button();
            this.NewBranchButton = new System.Windows.Forms.Button();
            this.PushButton = new System.Windows.Forms.Button();
            this.PullButton = new System.Windows.Forms.Button();
            this.SyncButton = new System.Windows.Forms.Button();
            this.RefreshButton = new System.Windows.Forms.ToolStripButton();
            this.toolStrip1.SuspendLayout();
            this.tabControl1.SuspendLayout();
            this.tabPage1.SuspendLayout();
            this.tabPage2.SuspendLayout();
            this.tabPage3.SuspendLayout();
            this.tabPage4.SuspendLayout();
            this.groupBox1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.IncludedChangesGrid)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.ExcludedChangesGrid)).BeginInit();
            this.groupBox3.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.IncomingCommitsGrid)).BeginInit();
            this.groupBox4.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.OutgoingCommitsGrid)).BeginInit();
            this.groupBox5.SuspendLayout();
            this.groupBox6.SuspendLayout();
            this.groupBox7.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.UnpublishedBranchesGrid)).BeginInit();
            this.groupBox8.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.PublishedBranchesGrid)).BeginInit();
            this.SuspendLayout();
            // 
            // toolStrip1
            // 
            this.toolStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.RefreshButton});
            this.toolStrip1.Location = new System.Drawing.Point(0, 0);
            this.toolStrip1.Name = "toolStrip1";
            this.toolStrip1.Size = new System.Drawing.Size(255, 25);
            this.toolStrip1.TabIndex = 0;
            this.toolStrip1.Text = "toolStrip1";
            // 
            // tabControl1
            // 
            this.tabControl1.Controls.Add(this.tabPage1);
            this.tabControl1.Controls.Add(this.tabPage2);
            this.tabControl1.Controls.Add(this.tabPage3);
            this.tabControl1.Controls.Add(this.tabPage4);
            this.tabControl1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tabControl1.Location = new System.Drawing.Point(0, 25);
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new System.Drawing.Size(255, 430);
            this.tabControl1.TabIndex = 1;
            // 
            // tabPage1
            // 
            this.tabPage1.Controls.Add(this.groupBox2);
            this.tabPage1.Controls.Add(this.groupBox1);
            this.tabPage1.Controls.Add(this.CommitButton);
            this.tabPage1.Controls.Add(this.CommitAction);
            this.tabPage1.Controls.Add(this.label2);
            this.tabPage1.Controls.Add(this.CommitMessage);
            this.tabPage1.Controls.Add(this.CurrentBranch);
            this.tabPage1.Controls.Add(this.label1);
            this.tabPage1.Location = new System.Drawing.Point(4, 22);
            this.tabPage1.Name = "tabPage1";
            this.tabPage1.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage1.Size = new System.Drawing.Size(247, 404);
            this.tabPage1.TabIndex = 0;
            this.tabPage1.Text = "Changes";
            this.tabPage1.UseVisualStyleBackColor = true;
            // 
            // tabPage2
            // 
            this.tabPage2.Controls.Add(this.groupBox7);
            this.tabPage2.Controls.Add(this.groupBox8);
            this.tabPage2.Controls.Add(this.comboBox2);
            this.tabPage2.Controls.Add(this.label4);
            this.tabPage2.Controls.Add(this.MergeButton);
            this.tabPage2.Controls.Add(this.NewBranchButton);
            this.tabPage2.Location = new System.Drawing.Point(4, 22);
            this.tabPage2.Name = "tabPage2";
            this.tabPage2.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage2.Size = new System.Drawing.Size(247, 404);
            this.tabPage2.TabIndex = 1;
            this.tabPage2.Text = "Branches";
            this.tabPage2.UseVisualStyleBackColor = true;
            // 
            // tabPage3
            // 
            this.tabPage3.Controls.Add(this.groupBox4);
            this.tabPage3.Controls.Add(this.groupBox3);
            this.tabPage3.Controls.Add(this.PushButton);
            this.tabPage3.Controls.Add(this.PullButton);
            this.tabPage3.Controls.Add(this.SyncButton);
            this.tabPage3.Controls.Add(this.comboBox1);
            this.tabPage3.Controls.Add(this.label3);
            this.tabPage3.Location = new System.Drawing.Point(4, 22);
            this.tabPage3.Name = "tabPage3";
            this.tabPage3.Size = new System.Drawing.Size(247, 404);
            this.tabPage3.TabIndex = 2;
            this.tabPage3.Text = "Unsynced commits";
            this.tabPage3.UseVisualStyleBackColor = true;
            // 
            // tabPage4
            // 
            this.tabPage4.Controls.Add(this.groupBox6);
            this.tabPage4.Controls.Add(this.groupBox5);
            this.tabPage4.Location = new System.Drawing.Point(4, 22);
            this.tabPage4.Name = "tabPage4";
            this.tabPage4.Size = new System.Drawing.Size(247, 404);
            this.tabPage4.TabIndex = 3;
            this.tabPage4.Text = "Settings";
            this.tabPage4.UseVisualStyleBackColor = true;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(6, 12);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(44, 13);
            this.label1.TabIndex = 0;
            this.label1.Text = "Branch:";
            // 
            // CurrentBranch
            // 
            this.CurrentBranch.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.CurrentBranch.FormattingEnabled = true;
            this.CurrentBranch.Location = new System.Drawing.Point(56, 9);
            this.CurrentBranch.Name = "CurrentBranch";
            this.CurrentBranch.Size = new System.Drawing.Size(185, 21);
            this.CurrentBranch.TabIndex = 1;
            // 
            // CommitMessage
            // 
            this.CommitMessage.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.CommitMessage.Location = new System.Drawing.Point(9, 57);
            this.CommitMessage.Multiline = true;
            this.CommitMessage.Name = "CommitMessage";
            this.CommitMessage.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.CommitMessage.Size = new System.Drawing.Size(232, 58);
            this.CommitMessage.TabIndex = 2;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(6, 41);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(89, 13);
            this.label2.TabIndex = 3;
            this.label2.Text = "Commit message:";
            // 
            // CommitAction
            // 
            this.CommitAction.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.CommitAction.FormattingEnabled = true;
            this.CommitAction.Items.AddRange(new object[] {
            "Commit",
            "Commit and Push",
            "Commit and Sync"});
            this.CommitAction.Location = new System.Drawing.Point(9, 122);
            this.CommitAction.Name = "CommitAction";
            this.CommitAction.Size = new System.Drawing.Size(121, 21);
            this.CommitAction.TabIndex = 4;
            // 
            // CommitButton
            // 
            this.CommitButton.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.CommitButton.Image = global::Rubberduck.Properties.Resources.tick;
            this.CommitButton.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.CommitButton.Location = new System.Drawing.Point(137, 122);
            this.CommitButton.Name = "CommitButton";
            this.CommitButton.Size = new System.Drawing.Size(104, 23);
            this.CommitButton.TabIndex = 5;
            this.CommitButton.Text = "Go";
            this.CommitButton.UseVisualStyleBackColor = true;
            // 
            // groupBox1
            // 
            this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.groupBox1.Controls.Add(this.IncludedChangesGrid);
            this.groupBox1.Location = new System.Drawing.Point(9, 163);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(232, 113);
            this.groupBox1.TabIndex = 6;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Included changes";
            // 
            // groupBox2
            // 
            this.groupBox2.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.groupBox2.Controls.Add(this.ExcludedChangesGrid);
            this.groupBox2.Location = new System.Drawing.Point(9, 282);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(232, 113);
            this.groupBox2.TabIndex = 7;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Excluded changes";
            // 
            // IncludedChangesGrid
            // 
            this.IncludedChangesGrid.AllowUserToAddRows = false;
            this.IncludedChangesGrid.AllowUserToDeleteRows = false;
            this.IncludedChangesGrid.AllowUserToResizeColumns = false;
            this.IncludedChangesGrid.AllowUserToResizeRows = false;
            this.IncludedChangesGrid.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.IncludedChangesGrid.BackgroundColor = System.Drawing.Color.White;
            this.IncludedChangesGrid.CellBorderStyle = System.Windows.Forms.DataGridViewCellBorderStyle.None;
            this.IncludedChangesGrid.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.IncludedChangesGrid.EditMode = System.Windows.Forms.DataGridViewEditMode.EditProgrammatically;
            this.IncludedChangesGrid.GridColor = System.Drawing.Color.White;
            this.IncludedChangesGrid.Location = new System.Drawing.Point(7, 20);
            this.IncludedChangesGrid.Name = "IncludedChangesGrid";
            this.IncludedChangesGrid.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.IncludedChangesGrid.Size = new System.Drawing.Size(219, 87);
            this.IncludedChangesGrid.TabIndex = 0;
            // 
            // ExcludedChangesGrid
            // 
            this.ExcludedChangesGrid.AllowUserToAddRows = false;
            this.ExcludedChangesGrid.AllowUserToDeleteRows = false;
            this.ExcludedChangesGrid.AllowUserToResizeColumns = false;
            this.ExcludedChangesGrid.AllowUserToResizeRows = false;
            this.ExcludedChangesGrid.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.ExcludedChangesGrid.BackgroundColor = System.Drawing.Color.White;
            this.ExcludedChangesGrid.CellBorderStyle = System.Windows.Forms.DataGridViewCellBorderStyle.None;
            this.ExcludedChangesGrid.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.ExcludedChangesGrid.EditMode = System.Windows.Forms.DataGridViewEditMode.EditProgrammatically;
            this.ExcludedChangesGrid.GridColor = System.Drawing.Color.White;
            this.ExcludedChangesGrid.Location = new System.Drawing.Point(7, 19);
            this.ExcludedChangesGrid.Name = "ExcludedChangesGrid";
            this.ExcludedChangesGrid.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.ExcludedChangesGrid.Size = new System.Drawing.Size(219, 87);
            this.ExcludedChangesGrid.TabIndex = 1;
            // 
            // comboBox1
            // 
            this.comboBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.comboBox1.FormattingEnabled = true;
            this.comboBox1.Location = new System.Drawing.Point(56, 9);
            this.comboBox1.Name = "comboBox1";
            this.comboBox1.Size = new System.Drawing.Size(185, 21);
            this.comboBox1.TabIndex = 3;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(6, 12);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(44, 13);
            this.label3.TabIndex = 2;
            this.label3.Text = "Branch:";
            // 
            // groupBox3
            // 
            this.groupBox3.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.groupBox3.Controls.Add(this.IncomingCommitsGrid);
            this.groupBox3.Location = new System.Drawing.Point(9, 77);
            this.groupBox3.Name = "groupBox3";
            this.groupBox3.Size = new System.Drawing.Size(232, 142);
            this.groupBox3.TabIndex = 7;
            this.groupBox3.TabStop = false;
            this.groupBox3.Text = "Incoming Commits";
            // 
            // IncomingCommitsGrid
            // 
            this.IncomingCommitsGrid.AllowUserToAddRows = false;
            this.IncomingCommitsGrid.AllowUserToDeleteRows = false;
            this.IncomingCommitsGrid.AllowUserToResizeColumns = false;
            this.IncomingCommitsGrid.AllowUserToResizeRows = false;
            this.IncomingCommitsGrid.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.IncomingCommitsGrid.BackgroundColor = System.Drawing.Color.White;
            this.IncomingCommitsGrid.CellBorderStyle = System.Windows.Forms.DataGridViewCellBorderStyle.None;
            this.IncomingCommitsGrid.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.IncomingCommitsGrid.EditMode = System.Windows.Forms.DataGridViewEditMode.EditProgrammatically;
            this.IncomingCommitsGrid.GridColor = System.Drawing.Color.White;
            this.IncomingCommitsGrid.Location = new System.Drawing.Point(7, 20);
            this.IncomingCommitsGrid.Name = "IncomingCommitsGrid";
            this.IncomingCommitsGrid.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.IncomingCommitsGrid.Size = new System.Drawing.Size(219, 116);
            this.IncomingCommitsGrid.TabIndex = 0;
            // 
            // groupBox4
            // 
            this.groupBox4.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.groupBox4.Controls.Add(this.OutgoingCommitsGrid);
            this.groupBox4.Location = new System.Drawing.Point(9, 225);
            this.groupBox4.Name = "groupBox4";
            this.groupBox4.Size = new System.Drawing.Size(232, 161);
            this.groupBox4.TabIndex = 8;
            this.groupBox4.TabStop = false;
            this.groupBox4.Text = "Outgoing Commits";
            // 
            // OutgoingCommitsGrid
            // 
            this.OutgoingCommitsGrid.AllowUserToAddRows = false;
            this.OutgoingCommitsGrid.AllowUserToDeleteRows = false;
            this.OutgoingCommitsGrid.AllowUserToResizeColumns = false;
            this.OutgoingCommitsGrid.AllowUserToResizeRows = false;
            this.OutgoingCommitsGrid.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.OutgoingCommitsGrid.BackgroundColor = System.Drawing.Color.White;
            this.OutgoingCommitsGrid.CellBorderStyle = System.Windows.Forms.DataGridViewCellBorderStyle.None;
            this.OutgoingCommitsGrid.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.OutgoingCommitsGrid.EditMode = System.Windows.Forms.DataGridViewEditMode.EditProgrammatically;
            this.OutgoingCommitsGrid.GridColor = System.Drawing.Color.White;
            this.OutgoingCommitsGrid.Location = new System.Drawing.Point(7, 20);
            this.OutgoingCommitsGrid.Name = "OutgoingCommitsGrid";
            this.OutgoingCommitsGrid.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.OutgoingCommitsGrid.Size = new System.Drawing.Size(219, 135);
            this.OutgoingCommitsGrid.TabIndex = 0;
            // 
            // comboBox2
            // 
            this.comboBox2.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.comboBox2.FormattingEnabled = true;
            this.comboBox2.Location = new System.Drawing.Point(56, 9);
            this.comboBox2.Name = "comboBox2";
            this.comboBox2.Size = new System.Drawing.Size(185, 21);
            this.comboBox2.TabIndex = 5;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(6, 12);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(44, 13);
            this.label4.TabIndex = 4;
            this.label4.Text = "Branch:";
            // 
            // groupBox5
            // 
            this.groupBox5.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.groupBox5.Controls.Add(this.CancelGlobalSettingsButton);
            this.groupBox5.Controls.Add(this.UpdateGlobalSettingsButton);
            this.groupBox5.Controls.Add(this.BrowseDefaultRepositoryLocationButton);
            this.groupBox5.Controls.Add(this.DefaultRepositoryLocation);
            this.groupBox5.Controls.Add(this.label7);
            this.groupBox5.Controls.Add(this.EmailAddress);
            this.groupBox5.Controls.Add(this.label6);
            this.groupBox5.Controls.Add(this.UserName);
            this.groupBox5.Controls.Add(this.label5);
            this.groupBox5.Location = new System.Drawing.Point(4, 4);
            this.groupBox5.Name = "groupBox5";
            this.groupBox5.Size = new System.Drawing.Size(240, 199);
            this.groupBox5.TabIndex = 0;
            this.groupBox5.TabStop = false;
            this.groupBox5.Text = "Global Settings";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(7, 20);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(60, 13);
            this.label5.TabIndex = 0;
            this.label5.Text = "User Name";
            // 
            // UserName
            // 
            this.UserName.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.UserName.Location = new System.Drawing.Point(10, 37);
            this.UserName.Name = "UserName";
            this.UserName.Size = new System.Drawing.Size(224, 20);
            this.UserName.TabIndex = 1;
            // 
            // EmailAddress
            // 
            this.EmailAddress.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.EmailAddress.Location = new System.Drawing.Point(9, 77);
            this.EmailAddress.Name = "EmailAddress";
            this.EmailAddress.Size = new System.Drawing.Size(224, 20);
            this.EmailAddress.TabIndex = 3;
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(6, 60);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(73, 13);
            this.label6.TabIndex = 2;
            this.label6.Text = "Email Address";
            // 
            // DefaultRepositoryLocation
            // 
            this.DefaultRepositoryLocation.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.DefaultRepositoryLocation.Location = new System.Drawing.Point(9, 117);
            this.DefaultRepositoryLocation.Name = "DefaultRepositoryLocation";
            this.DefaultRepositoryLocation.Size = new System.Drawing.Size(191, 20);
            this.DefaultRepositoryLocation.TabIndex = 5;
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(6, 100);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(138, 13);
            this.label7.TabIndex = 4;
            this.label7.Text = "Default Repository Location";
            // 
            // BrowseDefaultRepositoryLocationButton
            // 
            this.BrowseDefaultRepositoryLocationButton.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.BrowseDefaultRepositoryLocationButton.Location = new System.Drawing.Point(199, 117);
            this.BrowseDefaultRepositoryLocationButton.Name = "BrowseDefaultRepositoryLocationButton";
            this.BrowseDefaultRepositoryLocationButton.Size = new System.Drawing.Size(33, 20);
            this.BrowseDefaultRepositoryLocationButton.TabIndex = 6;
            this.BrowseDefaultRepositoryLocationButton.Text = "...";
            this.BrowseDefaultRepositoryLocationButton.UseVisualStyleBackColor = true;
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
            // CancelGlobalSettingsButton
            // 
            this.CancelGlobalSettingsButton.Location = new System.Drawing.Point(105, 168);
            this.CancelGlobalSettingsButton.Name = "CancelGlobalSettingsButton";
            this.CancelGlobalSettingsButton.Size = new System.Drawing.Size(92, 23);
            this.CancelGlobalSettingsButton.TabIndex = 8;
            this.CancelGlobalSettingsButton.Text = "Cancel";
            this.CancelGlobalSettingsButton.UseVisualStyleBackColor = true;
            // 
            // groupBox6
            // 
            this.groupBox6.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.groupBox6.Controls.Add(this.button2);
            this.groupBox6.Controls.Add(this.button1);
            this.groupBox6.Location = new System.Drawing.Point(4, 210);
            this.groupBox6.Name = "groupBox6";
            this.groupBox6.Size = new System.Drawing.Size(240, 100);
            this.groupBox6.TabIndex = 1;
            this.groupBox6.TabStop = false;
            this.groupBox6.Text = "Repository Settings";
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
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(105, 31);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(92, 23);
            this.button2.TabIndex = 1;
            this.button2.Text = "Attributes File";
            this.button2.UseVisualStyleBackColor = true;
            // 
            // groupBox7
            // 
            this.groupBox7.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.groupBox7.Controls.Add(this.UnpublishedBranchesGrid);
            this.groupBox7.Location = new System.Drawing.Point(6, 226);
            this.groupBox7.Name = "groupBox7";
            this.groupBox7.Size = new System.Drawing.Size(232, 161);
            this.groupBox7.TabIndex = 10;
            this.groupBox7.TabStop = false;
            this.groupBox7.Text = "Unpublished Branches";
            // 
            // UnpublishedBranchesGrid
            // 
            this.UnpublishedBranchesGrid.AllowUserToAddRows = false;
            this.UnpublishedBranchesGrid.AllowUserToDeleteRows = false;
            this.UnpublishedBranchesGrid.AllowUserToResizeColumns = false;
            this.UnpublishedBranchesGrid.AllowUserToResizeRows = false;
            this.UnpublishedBranchesGrid.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.UnpublishedBranchesGrid.BackgroundColor = System.Drawing.Color.White;
            this.UnpublishedBranchesGrid.CellBorderStyle = System.Windows.Forms.DataGridViewCellBorderStyle.None;
            this.UnpublishedBranchesGrid.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.UnpublishedBranchesGrid.EditMode = System.Windows.Forms.DataGridViewEditMode.EditProgrammatically;
            this.UnpublishedBranchesGrid.GridColor = System.Drawing.Color.White;
            this.UnpublishedBranchesGrid.Location = new System.Drawing.Point(7, 20);
            this.UnpublishedBranchesGrid.Name = "UnpublishedBranchesGrid";
            this.UnpublishedBranchesGrid.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.UnpublishedBranchesGrid.Size = new System.Drawing.Size(219, 135);
            this.UnpublishedBranchesGrid.TabIndex = 0;
            // 
            // groupBox8
            // 
            this.groupBox8.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.groupBox8.Controls.Add(this.PublishedBranchesGrid);
            this.groupBox8.Location = new System.Drawing.Point(6, 78);
            this.groupBox8.Name = "groupBox8";
            this.groupBox8.Size = new System.Drawing.Size(232, 142);
            this.groupBox8.TabIndex = 9;
            this.groupBox8.TabStop = false;
            this.groupBox8.Text = "Published Branches";
            // 
            // PublishedBranchesGrid
            // 
            this.PublishedBranchesGrid.AllowUserToAddRows = false;
            this.PublishedBranchesGrid.AllowUserToDeleteRows = false;
            this.PublishedBranchesGrid.AllowUserToResizeColumns = false;
            this.PublishedBranchesGrid.AllowUserToResizeRows = false;
            this.PublishedBranchesGrid.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.PublishedBranchesGrid.BackgroundColor = System.Drawing.Color.White;
            this.PublishedBranchesGrid.CellBorderStyle = System.Windows.Forms.DataGridViewCellBorderStyle.None;
            this.PublishedBranchesGrid.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.PublishedBranchesGrid.EditMode = System.Windows.Forms.DataGridViewEditMode.EditProgrammatically;
            this.PublishedBranchesGrid.GridColor = System.Drawing.Color.White;
            this.PublishedBranchesGrid.Location = new System.Drawing.Point(7, 20);
            this.PublishedBranchesGrid.Name = "PublishedBranchesGrid";
            this.PublishedBranchesGrid.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.PublishedBranchesGrid.Size = new System.Drawing.Size(219, 116);
            this.PublishedBranchesGrid.TabIndex = 0;
            // 
            // MergeButton
            // 
            this.MergeButton.Image = global::Rubberduck.Properties.Resources.arrow_merge_090;
            this.MergeButton.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.MergeButton.Location = new System.Drawing.Point(115, 40);
            this.MergeButton.Name = "MergeButton";
            this.MergeButton.Size = new System.Drawing.Size(75, 23);
            this.MergeButton.TabIndex = 7;
            this.MergeButton.Text = "Merge...";
            this.MergeButton.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText;
            this.MergeButton.UseVisualStyleBackColor = true;
            // 
            // NewBranchButton
            // 
            this.NewBranchButton.Image = global::Rubberduck.Properties.Resources.arrow_split;
            this.NewBranchButton.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.NewBranchButton.Location = new System.Drawing.Point(9, 40);
            this.NewBranchButton.Name = "NewBranchButton";
            this.NewBranchButton.Size = new System.Drawing.Size(99, 23);
            this.NewBranchButton.TabIndex = 6;
            this.NewBranchButton.Text = "New Branch...";
            this.NewBranchButton.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText;
            this.NewBranchButton.UseVisualStyleBackColor = true;
            // 
            // PushButton
            // 
            this.PushButton.Image = global::Rubberduck.Properties.Resources.drive_upload;
            this.PushButton.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.PushButton.Location = new System.Drawing.Point(160, 40);
            this.PushButton.Name = "PushButton";
            this.PushButton.Size = new System.Drawing.Size(63, 23);
            this.PushButton.TabIndex = 6;
            this.PushButton.Text = "Push";
            this.PushButton.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText;
            this.PushButton.UseVisualStyleBackColor = true;
            // 
            // PullButton
            // 
            this.PullButton.Image = global::Rubberduck.Properties.Resources.drive_download;
            this.PullButton.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.PullButton.Location = new System.Drawing.Point(91, 40);
            this.PullButton.Name = "PullButton";
            this.PullButton.Size = new System.Drawing.Size(63, 23);
            this.PullButton.TabIndex = 5;
            this.PullButton.Text = "Pull";
            this.PullButton.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText;
            this.PullButton.UseVisualStyleBackColor = true;
            this.PullButton.Click += new System.EventHandler(this.PullButton_Click);
            // 
            // SyncButton
            // 
            this.SyncButton.Image = global::Rubberduck.Properties.Resources.arrow_circle_double;
            this.SyncButton.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.SyncButton.Location = new System.Drawing.Point(9, 40);
            this.SyncButton.Name = "SyncButton";
            this.SyncButton.Size = new System.Drawing.Size(75, 23);
            this.SyncButton.TabIndex = 4;
            this.SyncButton.Text = "Sync";
            this.SyncButton.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText;
            this.SyncButton.UseVisualStyleBackColor = true;
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
            // 
            // SourceControlPanel
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.tabControl1);
            this.Controls.Add(this.toolStrip1);
            this.MinimumSize = new System.Drawing.Size(255, 455);
            this.Name = "SourceControlPanel";
            this.Size = new System.Drawing.Size(255, 455);
            this.toolStrip1.ResumeLayout(false);
            this.toolStrip1.PerformLayout();
            this.tabControl1.ResumeLayout(false);
            this.tabPage1.ResumeLayout(false);
            this.tabPage1.PerformLayout();
            this.tabPage2.ResumeLayout(false);
            this.tabPage2.PerformLayout();
            this.tabPage3.ResumeLayout(false);
            this.tabPage3.PerformLayout();
            this.tabPage4.ResumeLayout(false);
            this.groupBox1.ResumeLayout(false);
            this.groupBox2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.IncludedChangesGrid)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.ExcludedChangesGrid)).EndInit();
            this.groupBox3.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.IncomingCommitsGrid)).EndInit();
            this.groupBox4.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.OutgoingCommitsGrid)).EndInit();
            this.groupBox5.ResumeLayout(false);
            this.groupBox5.PerformLayout();
            this.groupBox6.ResumeLayout(false);
            this.groupBox7.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.UnpublishedBranchesGrid)).EndInit();
            this.groupBox8.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.PublishedBranchesGrid)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.ToolStrip toolStrip1;
        private System.Windows.Forms.ToolStripButton RefreshButton;
        private System.Windows.Forms.TabControl tabControl1;
        private System.Windows.Forms.TabPage tabPage1;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.DataGridView ExcludedChangesGrid;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.DataGridView IncludedChangesGrid;
        private System.Windows.Forms.Button CommitButton;
        private System.Windows.Forms.ComboBox CommitAction;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox CommitMessage;
        private System.Windows.Forms.ComboBox CurrentBranch;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TabPage tabPage2;
        private System.Windows.Forms.Button MergeButton;
        private System.Windows.Forms.Button NewBranchButton;
        private System.Windows.Forms.ComboBox comboBox2;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.TabPage tabPage3;
        private System.Windows.Forms.GroupBox groupBox4;
        private System.Windows.Forms.DataGridView OutgoingCommitsGrid;
        private System.Windows.Forms.GroupBox groupBox3;
        private System.Windows.Forms.DataGridView IncomingCommitsGrid;
        private System.Windows.Forms.Button PushButton;
        private System.Windows.Forms.Button PullButton;
        private System.Windows.Forms.Button SyncButton;
        private System.Windows.Forms.ComboBox comboBox1;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TabPage tabPage4;
        private System.Windows.Forms.GroupBox groupBox6;
        private System.Windows.Forms.GroupBox groupBox5;
        private System.Windows.Forms.Button CancelGlobalSettingsButton;
        private System.Windows.Forms.Button UpdateGlobalSettingsButton;
        private System.Windows.Forms.Button BrowseDefaultRepositoryLocationButton;
        private System.Windows.Forms.TextBox DefaultRepositoryLocation;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.TextBox EmailAddress;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.TextBox UserName;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.GroupBox groupBox7;
        private System.Windows.Forms.DataGridView UnpublishedBranchesGrid;
        private System.Windows.Forms.GroupBox groupBox8;
        private System.Windows.Forms.DataGridView PublishedBranchesGrid;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.Button button1;
    }
}
