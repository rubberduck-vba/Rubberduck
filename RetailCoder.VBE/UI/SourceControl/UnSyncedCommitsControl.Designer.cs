using System.ComponentModel;
using System.Windows.Forms;

namespace Rubberduck.UI.SourceControl
{
    partial class UnSyncedCommitsControl
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
            base.Dispose(disposing);
        }

        #region Component Designer generated code

        /// <summary> 
        /// Required method for Designer support - do not modify 
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.UnsyncedCommitsPanel = new System.Windows.Forms.Panel();
            this.UnsyncedCommitsBranchNameLabel = new System.Windows.Forms.Label();
            this.SyncButton = new System.Windows.Forms.Button();
            this.FetchIncomingCommitsButton = new System.Windows.Forms.Button();
            this.CurrentBranchLabel = new System.Windows.Forms.Label();
            this.PushButton = new System.Windows.Forms.Button();
            this.PullButton = new System.Windows.Forms.Button();
            this.OutgoingCommitsBox = new System.Windows.Forms.GroupBox();
            this.OutgoingCommitsGrid = new System.Windows.Forms.DataGridView();
            this.IncomingCommitsBox = new System.Windows.Forms.GroupBox();
            this.IncomingCommitsGrid = new System.Windows.Forms.DataGridView();
            this.UnsyncedCommitsPanel.SuspendLayout();
            this.OutgoingCommitsBox.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.OutgoingCommitsGrid)).BeginInit();
            this.IncomingCommitsBox.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.IncomingCommitsGrid)).BeginInit();
            this.SuspendLayout();
            // 
            // UnsyncedCommitsPanel
            // 
            this.UnsyncedCommitsPanel.AutoScroll = true;
            this.UnsyncedCommitsPanel.Controls.Add(this.UnsyncedCommitsBranchNameLabel);
            this.UnsyncedCommitsPanel.Controls.Add(this.SyncButton);
            this.UnsyncedCommitsPanel.Controls.Add(this.FetchIncomingCommitsButton);
            this.UnsyncedCommitsPanel.Controls.Add(this.CurrentBranchLabel);
            this.UnsyncedCommitsPanel.Controls.Add(this.PushButton);
            this.UnsyncedCommitsPanel.Controls.Add(this.PullButton);
            this.UnsyncedCommitsPanel.Controls.Add(this.OutgoingCommitsBox);
            this.UnsyncedCommitsPanel.Controls.Add(this.IncomingCommitsBox);
            this.UnsyncedCommitsPanel.Dock = System.Windows.Forms.DockStyle.Fill;
            this.UnsyncedCommitsPanel.Location = new System.Drawing.Point(0, 0);
            this.UnsyncedCommitsPanel.Name = "UnsyncedCommitsPanel";
            this.UnsyncedCommitsPanel.Padding = new System.Windows.Forms.Padding(3, 3, 3, 3);
            this.UnsyncedCommitsPanel.Size = new System.Drawing.Size(240, 482);
            this.UnsyncedCommitsPanel.TabIndex = 1;
            // 
            // UnsyncedCommitsBranchNameLabel
            // 
            this.UnsyncedCommitsBranchNameLabel.AutoSize = true;
            this.UnsyncedCommitsBranchNameLabel.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.UnsyncedCommitsBranchNameLabel.Location = new System.Drawing.Point(56, 14);
            this.UnsyncedCommitsBranchNameLabel.Name = "UnsyncedCommitsBranchNameLabel";
            this.UnsyncedCommitsBranchNameLabel.Size = new System.Drawing.Size(0, 13);
            this.UnsyncedCommitsBranchNameLabel.TabIndex = 19;
            // 
            // SyncButton
            // 
            this.SyncButton.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.SyncButton.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.SyncButton.Location = new System.Drawing.Point(9, 68);
            this.SyncButton.Name = "SyncButton";
            this.SyncButton.Size = new System.Drawing.Size(225, 23);
            this.SyncButton.TabIndex = 11;
            this.SyncButton.Text = "Sync";
            this.SyncButton.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText;
            this.SyncButton.UseVisualStyleBackColor = true;
            this.SyncButton.Click += new System.EventHandler(this.SyncButton_Click);
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
            this.FetchIncomingCommitsButton.Click += new System.EventHandler(this.FetchIncomingCommitsButton_Click);
            // 
            // CurrentBranchLabel
            // 
            this.CurrentBranchLabel.AutoSize = true;
            this.CurrentBranchLabel.Location = new System.Drawing.Point(6, 14);
            this.CurrentBranchLabel.Name = "CurrentBranchLabel";
            this.CurrentBranchLabel.Size = new System.Drawing.Size(44, 13);
            this.CurrentBranchLabel.TabIndex = 9;
            this.CurrentBranchLabel.Text = "Branch:";
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
            this.PushButton.Click += new System.EventHandler(this.PushButton_Click);
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
            this.PullButton.Click += new System.EventHandler(this.PullButton_Click);
            // 
            // OutgoingCommitsBox
            // 
            this.OutgoingCommitsBox.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.OutgoingCommitsBox.Controls.Add(this.OutgoingCommitsGrid);
            this.OutgoingCommitsBox.Location = new System.Drawing.Point(9, 269);
            this.OutgoingCommitsBox.Name = "OutgoingCommitsBox";
            this.OutgoingCommitsBox.Padding = new System.Windows.Forms.Padding(6, 6, 6, 6);
            this.OutgoingCommitsBox.Size = new System.Drawing.Size(225, 162);
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
            this.OutgoingCommitsGrid.Size = new System.Drawing.Size(206, 131);
            this.OutgoingCommitsGrid.TabIndex = 0;
            // 
            // IncomingCommitsBox
            // 
            this.IncomingCommitsBox.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.IncomingCommitsBox.Controls.Add(this.IncomingCommitsGrid);
            this.IncomingCommitsBox.Location = new System.Drawing.Point(9, 101);
            this.IncomingCommitsBox.Name = "IncomingCommitsBox";
            this.IncomingCommitsBox.Padding = new System.Windows.Forms.Padding(6, 6, 6, 6);
            this.IncomingCommitsBox.Size = new System.Drawing.Size(225, 162);
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
            this.IncomingCommitsGrid.Size = new System.Drawing.Size(206, 131);
            this.IncomingCommitsGrid.TabIndex = 0;
            // 
            // UnSyncedCommitsControl
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.AutoSize = true;
            this.Controls.Add(this.UnsyncedCommitsPanel);
            this.Name = "UnSyncedCommitsControl";
            this.Size = new System.Drawing.Size(240, 482);
            this.UnsyncedCommitsPanel.ResumeLayout(false);
            this.UnsyncedCommitsPanel.PerformLayout();
            this.OutgoingCommitsBox.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.OutgoingCommitsGrid)).EndInit();
            this.IncomingCommitsBox.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.IncomingCommitsGrid)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private Panel UnsyncedCommitsPanel;
        private Label UnsyncedCommitsBranchNameLabel;
        private Button SyncButton;
        private Button FetchIncomingCommitsButton;
        private Label CurrentBranchLabel;
        private Button PushButton;
        private Button PullButton;
        private GroupBox OutgoingCommitsBox;
        private DataGridView OutgoingCommitsGrid;
        private GroupBox IncomingCommitsBox;
        private DataGridView IncomingCommitsGrid;
    }
}
