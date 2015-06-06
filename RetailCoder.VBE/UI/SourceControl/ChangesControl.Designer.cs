using System.ComponentModel;
using System.Windows.Forms;

namespace Rubberduck.UI.SourceControl
{
    partial class ChangesControl
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
            this.ChangesPanel = new System.Windows.Forms.Panel();
            this.ChangesBranchNameLabel = new System.Windows.Forms.Label();
            this.IncludedChangesBox = new System.Windows.Forms.GroupBox();
            this.IncludedChangesGrid = new System.Windows.Forms.DataGridView();
            this.CommitButton = new System.Windows.Forms.Button();
            this.CommitActionDropdown = new System.Windows.Forms.ComboBox();
            this.CommitMessageLabel = new System.Windows.Forms.Label();
            this.CommitMessageBox = new System.Windows.Forms.TextBox();
            this.CurrentBranchLabel = new System.Windows.Forms.Label();
            this.ExcludedChangesBox = new System.Windows.Forms.GroupBox();
            this.ExcludedChangesGrid = new System.Windows.Forms.DataGridView();
            this.UntrackedFilesBox = new System.Windows.Forms.GroupBox();
            this.UntrackedFilesGrid = new System.Windows.Forms.DataGridView();
            this.ChangesPanel.SuspendLayout();
            this.IncludedChangesBox.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.IncludedChangesGrid)).BeginInit();
            this.ExcludedChangesBox.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.ExcludedChangesGrid)).BeginInit();
            this.UntrackedFilesBox.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.UntrackedFilesGrid)).BeginInit();
            this.SuspendLayout();
            // 
            // ChangesPanel
            // 
            this.ChangesPanel.AutoScroll = true;
            this.ChangesPanel.AutoSize = true;
            this.ChangesPanel.BackColor = System.Drawing.SystemColors.Control;
            this.ChangesPanel.Controls.Add(this.ChangesBranchNameLabel);
            this.ChangesPanel.Controls.Add(this.IncludedChangesBox);
            this.ChangesPanel.Controls.Add(this.CommitButton);
            this.ChangesPanel.Controls.Add(this.CommitActionDropdown);
            this.ChangesPanel.Controls.Add(this.CommitMessageLabel);
            this.ChangesPanel.Controls.Add(this.CommitMessageBox);
            this.ChangesPanel.Controls.Add(this.CurrentBranchLabel);
            this.ChangesPanel.Controls.Add(this.ExcludedChangesBox);
            this.ChangesPanel.Controls.Add(this.UntrackedFilesBox);
            this.ChangesPanel.Dock = System.Windows.Forms.DockStyle.Fill;
            this.ChangesPanel.Location = new System.Drawing.Point(0, 0);
            this.ChangesPanel.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.ChangesPanel.Name = "ChangesPanel";
            this.ChangesPanel.Padding = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.ChangesPanel.Size = new System.Drawing.Size(317, 690);
            this.ChangesPanel.TabIndex = 1;
            // 
            // ChangesBranchNameLabel
            // 
            this.ChangesBranchNameLabel.AutoSize = true;
            this.ChangesBranchNameLabel.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.ChangesBranchNameLabel.Location = new System.Drawing.Point(75, 17);
            this.ChangesBranchNameLabel.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.ChangesBranchNameLabel.Name = "ChangesBranchNameLabel";
            this.ChangesBranchNameLabel.Size = new System.Drawing.Size(57, 17);
            this.ChangesBranchNameLabel.TabIndex = 18;
            this.ChangesBranchNameLabel.Text = "Master";
            // 
            // IncludedChangesBox
            // 
            this.IncludedChangesBox.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.IncludedChangesBox.Controls.Add(this.IncludedChangesGrid);
            this.IncludedChangesBox.Location = new System.Drawing.Point(12, 146);
            this.IncludedChangesBox.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.IncludedChangesBox.Name = "IncludedChangesBox";
            this.IncludedChangesBox.Padding = new System.Windows.Forms.Padding(8, 7, 8, 7);
            this.IncludedChangesBox.Size = new System.Drawing.Size(297, 174);
            this.IncludedChangesBox.TabIndex = 15;
            this.IncludedChangesBox.TabStop = false;
            this.IncludedChangesBox.Text = "Included changes";
            // 
            // IncludedChangesGrid
            // 
            this.IncludedChangesGrid.AllowDrop = true;
            this.IncludedChangesGrid.AllowUserToAddRows = false;
            this.IncludedChangesGrid.AllowUserToDeleteRows = false;
            this.IncludedChangesGrid.AllowUserToResizeColumns = false;
            this.IncludedChangesGrid.AllowUserToResizeRows = false;
            this.IncludedChangesGrid.BackgroundColor = System.Drawing.Color.White;
            this.IncludedChangesGrid.CellBorderStyle = System.Windows.Forms.DataGridViewCellBorderStyle.None;
            this.IncludedChangesGrid.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.IncludedChangesGrid.Dock = System.Windows.Forms.DockStyle.Fill;
            this.IncludedChangesGrid.EditMode = System.Windows.Forms.DataGridViewEditMode.EditProgrammatically;
            this.IncludedChangesGrid.GridColor = System.Drawing.Color.White;
            this.IncludedChangesGrid.Location = new System.Drawing.Point(8, 22);
            this.IncludedChangesGrid.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.IncludedChangesGrid.Name = "IncludedChangesGrid";
            this.IncludedChangesGrid.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.IncludedChangesGrid.Size = new System.Drawing.Size(281, 145);
            this.IncludedChangesGrid.TabIndex = 0;
            this.IncludedChangesGrid.DragDrop += new System.Windows.Forms.DragEventHandler(this.IncludedChangesGrid_DragDrop);
            this.IncludedChangesGrid.DragOver += new System.Windows.Forms.DragEventHandler(this.IncludedChangesGrid_DragOver);
            this.IncludedChangesGrid.MouseDown += new System.Windows.Forms.MouseEventHandler(this.IncludedChangesGrid_MouseDown);
            this.IncludedChangesGrid.MouseMove += new System.Windows.Forms.MouseEventHandler(this.IncludedChangesGrid_MouseMove);
            // 
            // CommitButton
            // 
            this.CommitButton.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.CommitButton.AutoSize = true;
            this.CommitButton.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.CommitButton.Enabled = false;
            this.CommitButton.Image = global::Rubberduck.Properties.Resources.tick;
            this.CommitButton.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.CommitButton.Location = new System.Drawing.Point(227, 107);
            this.CommitButton.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.CommitButton.MinimumSize = new System.Drawing.Size(83, 28);
            this.CommitButton.Name = "CommitButton";
            this.CommitButton.Size = new System.Drawing.Size(83, 28);
            this.CommitButton.TabIndex = 14;
            this.CommitButton.Text = "Go";
            this.CommitButton.UseVisualStyleBackColor = true;
            this.CommitButton.Click += new System.EventHandler(this.CommitButton_Click);
            // 
            // CommitActionDropdown
            // 
            this.CommitActionDropdown.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.CommitActionDropdown.FormattingEnabled = true;
            this.CommitActionDropdown.Items.AddRange(new object[] {
            RubberduckUI.SourceControl_Commit,
            RubberduckUI.SourceControl_CommitPush,
            RubberduckUI.SourceControl_CommitSync});
            this.CommitActionDropdown.Location = new System.Drawing.Point(12, 107);
            this.CommitActionDropdown.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.CommitActionDropdown.MinimumSize = new System.Drawing.Size(160, 0);
            this.CommitActionDropdown.Name = "CommitActionDropdown";
            this.CommitActionDropdown.Size = new System.Drawing.Size(205, 24);
            this.CommitActionDropdown.TabIndex = 13;
            this.CommitActionDropdown.SelectedIndexChanged += new System.EventHandler(this.CommitActionDropdown_SelectedIndexChanged);
            // 
            // CommitMessageLabel
            // 
            this.CommitMessageLabel.AutoSize = true;
            this.CommitMessageLabel.Location = new System.Drawing.Point(8, 48);
            this.CommitMessageLabel.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.CommitMessageLabel.Name = "CommitMessageLabel";
            this.CommitMessageLabel.Size = new System.Drawing.Size(119, 17);
            this.CommitMessageLabel.TabIndex = 12;
            this.CommitMessageLabel.Text = "Commit message:";
            // 
            // CommitMessageBox
            // 
            this.CommitMessageBox.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.CommitMessageBox.BackColor = System.Drawing.Color.LightYellow;
            this.CommitMessageBox.Location = new System.Drawing.Point(12, 68);
            this.CommitMessageBox.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.CommitMessageBox.Multiline = true;
            this.CommitMessageBox.Name = "CommitMessageBox";
            this.CommitMessageBox.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.CommitMessageBox.Size = new System.Drawing.Size(300, 35);
            this.CommitMessageBox.TabIndex = 11;
            this.CommitMessageBox.TextChanged += new System.EventHandler(this.CommitMessageBox_TextChanged);
            // 
            // CurrentBranchLabel
            // 
            this.CurrentBranchLabel.AutoSize = true;
            this.CurrentBranchLabel.Location = new System.Drawing.Point(8, 17);
            this.CurrentBranchLabel.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.CurrentBranchLabel.Name = "CurrentBranchLabel";
            this.CurrentBranchLabel.Size = new System.Drawing.Size(57, 17);
            this.CurrentBranchLabel.TabIndex = 9;
            this.CurrentBranchLabel.Text = "Branch:";
            // 
            // ExcludedChangesBox
            // 
            this.ExcludedChangesBox.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.ExcludedChangesBox.Controls.Add(this.ExcludedChangesGrid);
            this.ExcludedChangesBox.Location = new System.Drawing.Point(12, 327);
            this.ExcludedChangesBox.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.ExcludedChangesBox.Name = "ExcludedChangesBox";
            this.ExcludedChangesBox.Padding = new System.Windows.Forms.Padding(8, 7, 8, 7);
            this.ExcludedChangesBox.Size = new System.Drawing.Size(297, 174);
            this.ExcludedChangesBox.TabIndex = 16;
            this.ExcludedChangesBox.TabStop = false;
            this.ExcludedChangesBox.Text = "Excluded changes";
            // 
            // ExcludedChangesGrid
            // 
            this.ExcludedChangesGrid.AllowDrop = true;
            this.ExcludedChangesGrid.AllowUserToAddRows = false;
            this.ExcludedChangesGrid.AllowUserToDeleteRows = false;
            this.ExcludedChangesGrid.AllowUserToResizeColumns = false;
            this.ExcludedChangesGrid.AllowUserToResizeRows = false;
            this.ExcludedChangesGrid.BackgroundColor = System.Drawing.Color.White;
            this.ExcludedChangesGrid.CellBorderStyle = System.Windows.Forms.DataGridViewCellBorderStyle.None;
            this.ExcludedChangesGrid.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.ExcludedChangesGrid.Dock = System.Windows.Forms.DockStyle.Fill;
            this.ExcludedChangesGrid.EditMode = System.Windows.Forms.DataGridViewEditMode.EditProgrammatically;
            this.ExcludedChangesGrid.GridColor = System.Drawing.Color.White;
            this.ExcludedChangesGrid.Location = new System.Drawing.Point(8, 22);
            this.ExcludedChangesGrid.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.ExcludedChangesGrid.Name = "ExcludedChangesGrid";
            this.ExcludedChangesGrid.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.ExcludedChangesGrid.Size = new System.Drawing.Size(281, 145);
            this.ExcludedChangesGrid.TabIndex = 1;
            this.ExcludedChangesGrid.DragDrop += new System.Windows.Forms.DragEventHandler(this.ExcludedChangesGrid_DragDrop);
            this.ExcludedChangesGrid.DragOver += new System.Windows.Forms.DragEventHandler(this.ExcludedChangesGrid_DragOver);
            this.ExcludedChangesGrid.MouseDown += new System.Windows.Forms.MouseEventHandler(this.ExcludedChangesGrid_MouseDown);
            this.ExcludedChangesGrid.MouseMove += new System.Windows.Forms.MouseEventHandler(this.ExcludedChangesGrid_MouseMove);
            // 
            // UntrackedFilesBox
            // 
            this.UntrackedFilesBox.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.UntrackedFilesBox.Controls.Add(this.UntrackedFilesGrid);
            this.UntrackedFilesBox.Location = new System.Drawing.Point(12, 508);
            this.UntrackedFilesBox.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.UntrackedFilesBox.Name = "UntrackedFilesBox";
            this.UntrackedFilesBox.Padding = new System.Windows.Forms.Padding(8, 7, 8, 7);
            this.UntrackedFilesBox.Size = new System.Drawing.Size(297, 174);
            this.UntrackedFilesBox.TabIndex = 17;
            this.UntrackedFilesBox.TabStop = false;
            this.UntrackedFilesBox.Text = "Untracked files";
            // 
            // UntrackedFilesGrid
            // 
            this.UntrackedFilesGrid.AllowUserToAddRows = false;
            this.UntrackedFilesGrid.AllowUserToDeleteRows = false;
            this.UntrackedFilesGrid.AllowUserToResizeColumns = false;
            this.UntrackedFilesGrid.AllowUserToResizeRows = false;
            this.UntrackedFilesGrid.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.UntrackedFilesGrid.BackgroundColor = System.Drawing.Color.White;
            this.UntrackedFilesGrid.CellBorderStyle = System.Windows.Forms.DataGridViewCellBorderStyle.None;
            this.UntrackedFilesGrid.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.UntrackedFilesGrid.EditMode = System.Windows.Forms.DataGridViewEditMode.EditProgrammatically;
            this.UntrackedFilesGrid.GridColor = System.Drawing.Color.White;
            this.UntrackedFilesGrid.Location = new System.Drawing.Point(13, 27);
            this.UntrackedFilesGrid.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.UntrackedFilesGrid.Name = "UntrackedFilesGrid";
            this.UntrackedFilesGrid.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.UntrackedFilesGrid.Size = new System.Drawing.Size(272, 135);
            this.UntrackedFilesGrid.TabIndex = 1;
            this.UntrackedFilesGrid.MouseDown += new System.Windows.Forms.MouseEventHandler(this.UntrackedFilesGrid_MouseDown);
            this.UntrackedFilesGrid.MouseMove += new System.Windows.Forms.MouseEventHandler(this.UntrackedFilesGrid_MouseMove);
            // 
            // ChangesControl
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.AutoSize = true;
            this.BackColor = System.Drawing.SystemColors.ControlLight;
            this.Controls.Add(this.ChangesPanel);
            this.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.Name = "ChangesControl";
            this.Size = new System.Drawing.Size(317, 690);
            this.ChangesPanel.ResumeLayout(false);
            this.ChangesPanel.PerformLayout();
            this.IncludedChangesBox.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.IncludedChangesGrid)).EndInit();
            this.ExcludedChangesBox.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.ExcludedChangesGrid)).EndInit();
            this.UntrackedFilesBox.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.UntrackedFilesGrid)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private Panel ChangesPanel;
        private Label ChangesBranchNameLabel;
        private GroupBox IncludedChangesBox;
        private DataGridView IncludedChangesGrid;
        private Button CommitButton;
        private ComboBox CommitActionDropdown;
        private Label CommitMessageLabel;
        private TextBox CommitMessageBox;
        private Label CurrentBranchLabel;
        private GroupBox ExcludedChangesBox;
        private DataGridView ExcludedChangesGrid;
        private GroupBox UntrackedFilesBox;
        private DataGridView UntrackedFilesGrid;
    }
}
