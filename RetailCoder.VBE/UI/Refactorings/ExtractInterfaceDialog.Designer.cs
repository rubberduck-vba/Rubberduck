namespace Rubberduck.UI.Refactorings
{
    partial class ExtractInterfaceDialog
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

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(ExtractInterfaceDialog));
            this.flowLayoutPanel2 = new System.Windows.Forms.FlowLayoutPanel();
            this.CancelDialogButton = new System.Windows.Forms.Button();
            this.OkButton = new System.Windows.Forms.Button();
            this.InvalidNameValidationIcon = new System.Windows.Forms.PictureBox();
            this.InterfaceNameBox = new System.Windows.Forms.TextBox();
            this.NameLabel = new System.Windows.Forms.Label();
            this.TitleLabel = new System.Windows.Forms.Label();
            this.InstructionsLabel = new System.Windows.Forms.Label();
            this.DescriptionPanel = new System.Windows.Forms.Panel();
            this.MembersGroupBox = new System.Windows.Forms.GroupBox();
            this.InterfaceMembersGridView = new System.Windows.Forms.DataGridView();
            this.DeselectAllButton = new System.Windows.Forms.Button();
            this.SelectAllButton = new System.Windows.Forms.Button();
            this.flowLayoutPanel2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.InvalidNameValidationIcon)).BeginInit();
            this.DescriptionPanel.SuspendLayout();
            this.MembersGroupBox.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.InterfaceMembersGridView)).BeginInit();
            this.SuspendLayout();
            // 
            // flowLayoutPanel2
            // 
            this.flowLayoutPanel2.BackColor = System.Drawing.SystemColors.ControlDark;
            this.flowLayoutPanel2.Controls.Add(this.CancelDialogButton);
            this.flowLayoutPanel2.Controls.Add(this.OkButton);
            this.flowLayoutPanel2.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.flowLayoutPanel2.FlowDirection = System.Windows.Forms.FlowDirection.RightToLeft;
            this.flowLayoutPanel2.Location = new System.Drawing.Point(0, 364);
            this.flowLayoutPanel2.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.flowLayoutPanel2.Name = "flowLayoutPanel2";
            this.flowLayoutPanel2.Padding = new System.Windows.Forms.Padding(11, 10, 0, 10);
            this.flowLayoutPanel2.Size = new System.Drawing.Size(572, 53);
            this.flowLayoutPanel2.TabIndex = 28;
            // 
            // CancelDialogButton
            // 
            this.CancelDialogButton.BackColor = System.Drawing.SystemColors.ButtonFace;
            this.CancelDialogButton.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.CancelDialogButton.Location = new System.Drawing.Point(457, 14);
            this.CancelDialogButton.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.CancelDialogButton.Name = "CancelDialogButton";
            this.CancelDialogButton.Size = new System.Drawing.Size(100, 28);
            this.CancelDialogButton.TabIndex = 0;
            this.CancelDialogButton.Text = "Cancel";
            this.CancelDialogButton.UseVisualStyleBackColor = false;
            // 
            // OkButton
            // 
            this.OkButton.BackColor = System.Drawing.SystemColors.ButtonFace;
            this.OkButton.DialogResult = System.Windows.Forms.DialogResult.OK;
            this.OkButton.Location = new System.Drawing.Point(349, 14);
            this.OkButton.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.OkButton.Name = "OkButton";
            this.OkButton.Size = new System.Drawing.Size(100, 28);
            this.OkButton.TabIndex = 1;
            this.OkButton.Text = "Ok";
            this.OkButton.UseVisualStyleBackColor = false;
            // 
            // InvalidNameValidationIcon
            // 
            this.InvalidNameValidationIcon.Image = global::Rubberduck.Properties.Resources.cross_circle;
            this.InvalidNameValidationIcon.Location = new System.Drawing.Point(549, 102);
            this.InvalidNameValidationIcon.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.InvalidNameValidationIcon.Name = "InvalidNameValidationIcon";
            this.InvalidNameValidationIcon.Size = new System.Drawing.Size(16, 16);
            this.InvalidNameValidationIcon.SizeMode = System.Windows.Forms.PictureBoxSizeMode.AutoSize;
            this.InvalidNameValidationIcon.TabIndex = 32;
            this.InvalidNameValidationIcon.TabStop = false;
            // 
            // InterfaceNameBox
            // 
            this.InterfaceNameBox.Location = new System.Drawing.Point(16, 112);
            this.InterfaceNameBox.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.InterfaceNameBox.Name = "InterfaceNameBox";
            this.InterfaceNameBox.Size = new System.Drawing.Size(541, 22);
            this.InterfaceNameBox.TabIndex = 30;
            // 
            // NameLabel
            // 
            this.NameLabel.AutoSize = true;
            this.NameLabel.Location = new System.Drawing.Point(13, 91);
            this.NameLabel.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.NameLabel.Name = "NameLabel";
            this.NameLabel.Size = new System.Drawing.Size(49, 17);
            this.NameLabel.TabIndex = 29;
            this.NameLabel.Text = "Name:";
            // 
            // TitleLabel
            // 
            this.TitleLabel.AutoSize = true;
            this.TitleLabel.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold);
            this.TitleLabel.Location = new System.Drawing.Point(20, 11);
            this.TitleLabel.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.TitleLabel.Name = "TitleLabel";
            this.TitleLabel.Padding = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.TitleLabel.Size = new System.Drawing.Size(137, 22);
            this.TitleLabel.TabIndex = 2;
            this.TitleLabel.Text = "Extract Interface";
            // 
            // InstructionsLabel
            // 
            this.InstructionsLabel.Location = new System.Drawing.Point(20, 30);
            this.InstructionsLabel.Name = "InstructionsLabel";
            this.InstructionsLabel.Padding = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.InstructionsLabel.Size = new System.Drawing.Size(549, 34);
            this.InstructionsLabel.TabIndex = 3;
            this.InstructionsLabel.Text = "Please specify interface name and members.";
            // 
            // DescriptionPanel
            // 
            this.DescriptionPanel.BackColor = System.Drawing.Color.White;
            this.DescriptionPanel.Controls.Add(this.TitleLabel);
            this.DescriptionPanel.Controls.Add(this.InstructionsLabel);
            this.DescriptionPanel.Dock = System.Windows.Forms.DockStyle.Top;
            this.DescriptionPanel.Location = new System.Drawing.Point(0, 0);
            this.DescriptionPanel.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.DescriptionPanel.Name = "DescriptionPanel";
            this.DescriptionPanel.Size = new System.Drawing.Size(572, 84);
            this.DescriptionPanel.TabIndex = 33;
            // 
            // MembersGroupBox
            // 
            this.MembersGroupBox.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.MembersGroupBox.Controls.Add(this.InterfaceMembersGridView);
            this.MembersGroupBox.Controls.Add(this.DeselectAllButton);
            this.MembersGroupBox.Controls.Add(this.SelectAllButton);
            this.MembersGroupBox.Location = new System.Drawing.Point(16, 142);
            this.MembersGroupBox.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.MembersGroupBox.Name = "MembersGroupBox";
            this.MembersGroupBox.Padding = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.MembersGroupBox.Size = new System.Drawing.Size(541, 214);
            this.MembersGroupBox.TabIndex = 34;
            this.MembersGroupBox.TabStop = false;
            this.MembersGroupBox.Text = "Members";
            // 
            // InterfaceMembersGridView
            // 
            this.InterfaceMembersGridView.AllowUserToAddRows = false;
            this.InterfaceMembersGridView.AllowUserToDeleteRows = false;
            this.InterfaceMembersGridView.AllowUserToResizeColumns = false;
            this.InterfaceMembersGridView.AllowUserToResizeRows = false;
            this.InterfaceMembersGridView.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.InterfaceMembersGridView.ColumnHeadersVisible = false;
            this.InterfaceMembersGridView.EditMode = System.Windows.Forms.DataGridViewEditMode.EditOnEnter;
            this.InterfaceMembersGridView.Location = new System.Drawing.Point(7, 26);
            this.InterfaceMembersGridView.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.InterfaceMembersGridView.MultiSelect = false;
            this.InterfaceMembersGridView.Name = "InterfaceMembersGridView";
            this.InterfaceMembersGridView.RowHeadersVisible = false;
            this.InterfaceMembersGridView.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing;
            this.InterfaceMembersGridView.RowTemplate.Height = 24;
            this.InterfaceMembersGridView.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.InterfaceMembersGridView.ShowEditingIcon = false;
            this.InterfaceMembersGridView.Size = new System.Drawing.Size(427, 174);
            this.InterfaceMembersGridView.TabIndex = 3;
            // 
            // DeselectAllButton
            // 
            this.DeselectAllButton.Location = new System.Drawing.Point(441, 64);
            this.DeselectAllButton.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.DeselectAllButton.Name = "DeselectAllButton";
            this.DeselectAllButton.Size = new System.Drawing.Size(93, 32);
            this.DeselectAllButton.TabIndex = 1;
            this.DeselectAllButton.Text = "Deselect All";
            this.DeselectAllButton.UseVisualStyleBackColor = true;
            // 
            // SelectAllButton
            // 
            this.SelectAllButton.Location = new System.Drawing.Point(441, 26);
            this.SelectAllButton.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.SelectAllButton.Name = "SelectAllButton";
            this.SelectAllButton.Size = new System.Drawing.Size(93, 32);
            this.SelectAllButton.TabIndex = 0;
            this.SelectAllButton.Text = "Select All";
            this.SelectAllButton.UseVisualStyleBackColor = true;
            // 
            // ExtractInterfaceDialog
            // 
            this.AcceptButton = this.OkButton;
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.CancelButton = this.CancelDialogButton;
            this.ClientSize = new System.Drawing.Size(572, 417);
            this.Controls.Add(this.MembersGroupBox);
            this.Controls.Add(this.DescriptionPanel);
            this.Controls.Add(this.InvalidNameValidationIcon);
            this.Controls.Add(this.InterfaceNameBox);
            this.Controls.Add(this.NameLabel);
            this.Controls.Add(this.flowLayoutPanel2);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "ExtractInterfaceDialog";
            this.ShowInTaskbar = false;
            this.Text = "Rubberduck - Extract Interface";
            this.flowLayoutPanel2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.InvalidNameValidationIcon)).EndInit();
            this.DescriptionPanel.ResumeLayout(false);
            this.DescriptionPanel.PerformLayout();
            this.MembersGroupBox.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.InterfaceMembersGridView)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.FlowLayoutPanel flowLayoutPanel2;
        private System.Windows.Forms.Button CancelDialogButton;
        private System.Windows.Forms.Button OkButton;
        private System.Windows.Forms.PictureBox InvalidNameValidationIcon;
        private System.Windows.Forms.TextBox InterfaceNameBox;
        private System.Windows.Forms.Label NameLabel;
        private System.Windows.Forms.Label TitleLabel;
        private System.Windows.Forms.Label InstructionsLabel;
        private System.Windows.Forms.Panel DescriptionPanel;
        private System.Windows.Forms.GroupBox MembersGroupBox;
        private System.Windows.Forms.Button DeselectAllButton;
        private System.Windows.Forms.Button SelectAllButton;
        private System.Windows.Forms.DataGridView InterfaceMembersGridView;
    }
}
