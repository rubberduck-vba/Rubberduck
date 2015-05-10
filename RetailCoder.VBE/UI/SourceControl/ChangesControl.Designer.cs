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
            this.label2 = new System.Windows.Forms.Label();
            this.CommitMessageBox = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
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
            this.ChangesPanel.Controls.Add(this.ChangesBranchNameLabel);
            this.ChangesPanel.Controls.Add(this.IncludedChangesBox);
            this.ChangesPanel.Controls.Add(this.CommitButton);
            this.ChangesPanel.Controls.Add(this.CommitActionDropdown);
            this.ChangesPanel.Controls.Add(this.label2);
            this.ChangesPanel.Controls.Add(this.CommitMessageBox);
            this.ChangesPanel.Controls.Add(this.label1);
            this.ChangesPanel.Controls.Add(this.ExcludedChangesBox);
            this.ChangesPanel.Controls.Add(this.UntrackedFilesBox);
            this.ChangesPanel.Dock = System.Windows.Forms.DockStyle.Fill;
            this.ChangesPanel.Location = new System.Drawing.Point(0, 0);
            this.ChangesPanel.Name = "ChangesPanel";
            this.ChangesPanel.Padding = new System.Windows.Forms.Padding(3);
            this.ChangesPanel.Size = new System.Drawing.Size(238, 560);
            this.ChangesPanel.TabIndex = 1;
            // 
            // ChangesBranchNameLabel
            // 
            this.ChangesBranchNameLabel.AutoSize = true;
            this.ChangesBranchNameLabel.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.ChangesBranchNameLabel.Location = new System.Drawing.Point(56, 14);
            this.ChangesBranchNameLabel.Name = "ChangesBranchNameLabel";
            this.ChangesBranchNameLabel.Size = new System.Drawing.Size(45, 13);
            this.ChangesBranchNameLabel.TabIndex = 18;
            this.ChangesBranchNameLabel.Text = "Master";
            // 
            // IncludedChangesBox
            // 
            this.IncludedChangesBox.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.IncludedChangesBox.Controls.Add(this.IncludedChangesGrid);
            this.IncludedChangesBox.Location = new System.Drawing.Point(9, 119);
            this.IncludedChangesBox.Name = "IncludedChangesBox";
            this.IncludedChangesBox.Padding = new System.Windows.Forms.Padding(6);
            this.IncludedChangesBox.Size = new System.Drawing.Size(183, 141);
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
            this.IncludedChangesGrid.Location = new System.Drawing.Point(6, 19);
            this.IncludedChangesGrid.Name = "IncludedChangesGrid";
            this.IncludedChangesGrid.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.IncludedChangesGrid.Size = new System.Drawing.Size(171, 116);
            this.IncludedChangesGrid.TabIndex = 0;
            this.IncludedChangesGrid.DragDrop += new System.Windows.Forms.DragEventHandler(this.IncludedChangesGrid_DragDrop);
            this.IncludedChangesGrid.DragOver += new System.Windows.Forms.DragEventHandler(this.IncludedChangesGrid_DragOver);
            this.IncludedChangesGrid.MouseDown += new System.Windows.Forms.MouseEventHandler(this.IncludedChangesGrid_MouseDown);
            this.IncludedChangesGrid.MouseMove += new System.Windows.Forms.MouseEventHandler(this.IncludedChangesGrid_MouseMove);
            // 
            // CommitButton
            // 
            this.CommitButton.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.CommitButton.AutoSize = true;
            this.CommitButton.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.CommitButton.Enabled = false;
            this.CommitButton.Image = global::Rubberduck.Properties.Resources.tick;
            this.CommitButton.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.CommitButton.Location = new System.Drawing.Point(136, 86);
            this.CommitButton.MinimumSize = new System.Drawing.Size(62, 23);
            this.CommitButton.Name = "CommitButton";
            this.CommitButton.Size = new System.Drawing.Size(62, 23);
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
            "Commit",
            "Commit and Push",
            "Commit and Sync"});
            this.CommitActionDropdown.Location = new System.Drawing.Point(9, 87);
            this.CommitActionDropdown.MinimumSize = new System.Drawing.Size(121, 0);
            this.CommitActionDropdown.Name = "CommitActionDropdown";
            this.CommitActionDropdown.Size = new System.Drawing.Size(121, 21);
            this.CommitActionDropdown.TabIndex = 13;
            this.CommitActionDropdown.SelectedIndexChanged += new System.EventHandler(this.CommitActionDropdown_SelectedIndexChanged);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(6, 39);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(89, 13);
            this.label2.TabIndex = 12;
            this.label2.Text = "Commit message:";
            // 
            // CommitMessageBox
            // 
            this.CommitMessageBox.BackColor = System.Drawing.Color.LightYellow;
            this.CommitMessageBox.Location = new System.Drawing.Point(9, 55);
            this.CommitMessageBox.Multiline = true;
            this.CommitMessageBox.Name = "CommitMessageBox";
            this.CommitMessageBox.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.CommitMessageBox.Size = new System.Drawing.Size(189, 29);
            this.CommitMessageBox.TabIndex = 11;
            this.CommitMessageBox.TextChanged += new System.EventHandler(this.CommitMessageBox_TextChanged);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(6, 14);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(44, 13);
            this.label1.TabIndex = 9;
            this.label1.Text = "Branch:";
            // 
            // ExcludedChangesBox
            // 
            this.ExcludedChangesBox.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.ExcludedChangesBox.Controls.Add(this.ExcludedChangesGrid);
            this.ExcludedChangesBox.Location = new System.Drawing.Point(9, 266);
            this.ExcludedChangesBox.Name = "ExcludedChangesBox";
            this.ExcludedChangesBox.Padding = new System.Windows.Forms.Padding(6);
            this.ExcludedChangesBox.Size = new System.Drawing.Size(177, 141);
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
            this.ExcludedChangesGrid.Location = new System.Drawing.Point(6, 19);
            this.ExcludedChangesGrid.Name = "ExcludedChangesGrid";
            this.ExcludedChangesGrid.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.ExcludedChangesGrid.Size = new System.Drawing.Size(165, 116);
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
            this.UntrackedFilesBox.Location = new System.Drawing.Point(9, 413);
            this.UntrackedFilesBox.Name = "UntrackedFilesBox";
            this.UntrackedFilesBox.Padding = new System.Windows.Forms.Padding(6);
            this.UntrackedFilesBox.Size = new System.Drawing.Size(177, 141);
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
            this.UntrackedFilesGrid.Location = new System.Drawing.Point(10, 22);
            this.UntrackedFilesGrid.Name = "UntrackedFilesGrid";
            this.UntrackedFilesGrid.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.UntrackedFilesGrid.Size = new System.Drawing.Size(158, 110);
            this.UntrackedFilesGrid.TabIndex = 1;
            this.UntrackedFilesGrid.MouseDown += new System.Windows.Forms.MouseEventHandler(this.UntrackedFilesGrid_MouseDown);
            this.UntrackedFilesGrid.MouseMove += new System.Windows.Forms.MouseEventHandler(this.UntrackedFilesGrid_MouseMove);
            // 
            // ChangesControl
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.AutoSize = true;
            this.Controls.Add(this.ChangesPanel);
            this.Name = "ChangesControl";
            this.Size = new System.Drawing.Size(238, 560);
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
        private Label label2;
        private TextBox CommitMessageBox;
        private Label label1;
        private GroupBox ExcludedChangesBox;
        private DataGridView ExcludedChangesGrid;
        private GroupBox UntrackedFilesBox;
        private DataGridView UntrackedFilesGrid;
    }
}
