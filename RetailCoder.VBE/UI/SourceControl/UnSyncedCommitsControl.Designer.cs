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
            this.label3 = new System.Windows.Forms.Label();
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
            this.UnsyncedCommitsPanel.Controls.Add(this.label3);
            this.UnsyncedCommitsPanel.Controls.Add(this.PushButton);
            this.UnsyncedCommitsPanel.Controls.Add(this.PullButton);
            this.UnsyncedCommitsPanel.Controls.Add(this.OutgoingCommitsBox);
            this.UnsyncedCommitsPanel.Controls.Add(this.IncomingCommitsBox);
            this.UnsyncedCommitsPanel.Dock = System.Windows.Forms.DockStyle.Fill;
            this.UnsyncedCommitsPanel.Location = new System.Drawing.Point(0, 0);
            this.UnsyncedCommitsPanel.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.UnsyncedCommitsPanel.Name = "UnsyncedCommitsPanel";
            this.UnsyncedCommitsPanel.Padding = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.UnsyncedCommitsPanel.Size = new System.Drawing.Size(320, 593);
            this.UnsyncedCommitsPanel.TabIndex = 1;
            // 
            // UnsyncedCommitsBranchNameLabel
            // 
            this.UnsyncedCommitsBranchNameLabel.AutoSize = true;
            this.UnsyncedCommitsBranchNameLabel.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.UnsyncedCommitsBranchNameLabel.Location = new System.Drawing.Point(75, 17);
            this.UnsyncedCommitsBranchNameLabel.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.UnsyncedCommitsBranchNameLabel.Name = "UnsyncedCommitsBranchNameLabel";
            this.UnsyncedCommitsBranchNameLabel.Size = new System.Drawing.Size(57, 17);
            this.UnsyncedCommitsBranchNameLabel.TabIndex = 19;
            this.UnsyncedCommitsBranchNameLabel.Text = "Master";
            // 
            // SyncButton
            // 
            this.SyncButton.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.SyncButton.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.SyncButton.Location = new System.Drawing.Point(12, 84);
            this.SyncButton.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.SyncButton.Name = "SyncButton";
            this.SyncButton.Size = new System.Drawing.Size(300, 28);
            this.SyncButton.TabIndex = 11;
            this.SyncButton.Text = "Sync";
            this.SyncButton.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText;
            this.SyncButton.UseVisualStyleBackColor = true;
            // 
            // FetchIncomingCommitsButton
            // 
            this.FetchIncomingCommitsButton.Image = global::Rubberduck.Properties.Resources.arrow_step;
            this.FetchIncomingCommitsButton.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.FetchIncomingCommitsButton.Location = new System.Drawing.Point(12, 48);
            this.FetchIncomingCommitsButton.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.FetchIncomingCommitsButton.Name = "FetchIncomingCommitsButton";
            this.FetchIncomingCommitsButton.Size = new System.Drawing.Size(84, 28);
            this.FetchIncomingCommitsButton.TabIndex = 13;
            this.FetchIncomingCommitsButton.Text = "Fetch";
            this.FetchIncomingCommitsButton.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText;
            this.FetchIncomingCommitsButton.UseVisualStyleBackColor = true;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(8, 17);
            this.label3.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(57, 17);
            this.label3.TabIndex = 9;
            this.label3.Text = "Branch:";
            // 
            // PushButton
            // 
            this.PushButton.Image = global::Rubberduck.Properties.Resources.drive_upload;
            this.PushButton.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.PushButton.Location = new System.Drawing.Point(196, 48);
            this.PushButton.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.PushButton.Name = "PushButton";
            this.PushButton.Size = new System.Drawing.Size(84, 28);
            this.PushButton.TabIndex = 14;
            this.PushButton.Text = "Push";
            this.PushButton.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText;
            this.PushButton.UseVisualStyleBackColor = true;
            // 
            // PullButton
            // 
            this.PullButton.Image = global::Rubberduck.Properties.Resources.drive_download;
            this.PullButton.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.PullButton.Location = new System.Drawing.Point(104, 48);
            this.PullButton.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.PullButton.Name = "PullButton";
            this.PullButton.Size = new System.Drawing.Size(84, 28);
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
            this.OutgoingCommitsBox.Location = new System.Drawing.Point(12, 331);
            this.OutgoingCommitsBox.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.OutgoingCommitsBox.Name = "OutgoingCommitsBox";
            this.OutgoingCommitsBox.Padding = new System.Windows.Forms.Padding(8, 7, 8, 7);
            this.OutgoingCommitsBox.Size = new System.Drawing.Size(300, 199);
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
            this.OutgoingCommitsGrid.Location = new System.Drawing.Point(13, 27);
            this.OutgoingCommitsGrid.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.OutgoingCommitsGrid.Name = "OutgoingCommitsGrid";
            this.OutgoingCommitsGrid.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.OutgoingCommitsGrid.Size = new System.Drawing.Size(275, 161);
            this.OutgoingCommitsGrid.TabIndex = 0;
            // 
            // IncomingCommitsBox
            // 
            this.IncomingCommitsBox.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.IncomingCommitsBox.Controls.Add(this.IncomingCommitsGrid);
            this.IncomingCommitsBox.Location = new System.Drawing.Point(12, 124);
            this.IncomingCommitsBox.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.IncomingCommitsBox.Name = "IncomingCommitsBox";
            this.IncomingCommitsBox.Padding = new System.Windows.Forms.Padding(8, 7, 8, 7);
            this.IncomingCommitsBox.Size = new System.Drawing.Size(300, 199);
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
            this.IncomingCommitsGrid.Location = new System.Drawing.Point(13, 27);
            this.IncomingCommitsGrid.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.IncomingCommitsGrid.Name = "IncomingCommitsGrid";
            this.IncomingCommitsGrid.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.IncomingCommitsGrid.Size = new System.Drawing.Size(275, 161);
            this.IncomingCommitsGrid.TabIndex = 0;
            // 
            // UnSyncedCommitsControl
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.AutoSize = true;
            this.Controls.Add(this.UnsyncedCommitsPanel);
            this.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.Name = "UnSyncedCommitsControl";
            this.Size = new System.Drawing.Size(320, 593);
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
        private Label label3;
        private Button PushButton;
        private Button PullButton;
        private GroupBox OutgoingCommitsBox;
        private DataGridView OutgoingCommitsGrid;
        private GroupBox IncomingCommitsBox;
        private DataGridView IncomingCommitsGrid;
    }
}
