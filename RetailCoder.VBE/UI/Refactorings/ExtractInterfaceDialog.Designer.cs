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
            resources.ApplyResources(this.flowLayoutPanel2, "flowLayoutPanel2");
            this.flowLayoutPanel2.Name = "flowLayoutPanel2";
            // 
            // CancelDialogButton
            // 
            this.CancelDialogButton.BackColor = System.Drawing.SystemColors.ButtonFace;
            this.CancelDialogButton.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            resources.ApplyResources(this.CancelDialogButton, "CancelDialogButton");
            this.CancelDialogButton.Name = "CancelDialogButton";
            this.CancelDialogButton.UseVisualStyleBackColor = false;
            // 
            // OkButton
            // 
            this.OkButton.BackColor = System.Drawing.SystemColors.ButtonFace;
            this.OkButton.DialogResult = System.Windows.Forms.DialogResult.OK;
            resources.ApplyResources(this.OkButton, "OkButton");
            this.OkButton.Name = "OkButton";
            this.OkButton.UseVisualStyleBackColor = false;
            // 
            // InvalidNameValidationIcon
            // 
            this.InvalidNameValidationIcon.Image = global::Rubberduck.Properties.Resources.cross_circle;
            resources.ApplyResources(this.InvalidNameValidationIcon, "InvalidNameValidationIcon");
            this.InvalidNameValidationIcon.Name = "InvalidNameValidationIcon";
            this.InvalidNameValidationIcon.TabStop = false;
            // 
            // InterfaceNameBox
            // 
            resources.ApplyResources(this.InterfaceNameBox, "InterfaceNameBox");
            this.InterfaceNameBox.Name = "InterfaceNameBox";
            // 
            // NameLabel
            // 
            resources.ApplyResources(this.NameLabel, "NameLabel");
            this.NameLabel.Name = "NameLabel";
            // 
            // TitleLabel
            // 
            resources.ApplyResources(this.TitleLabel, "TitleLabel");
            this.TitleLabel.Name = "TitleLabel";
            // 
            // InstructionsLabel
            // 
            resources.ApplyResources(this.InstructionsLabel, "InstructionsLabel");
            this.InstructionsLabel.Name = "InstructionsLabel";
            // 
            // DescriptionPanel
            // 
            this.DescriptionPanel.BackColor = System.Drawing.Color.White;
            this.DescriptionPanel.Controls.Add(this.TitleLabel);
            this.DescriptionPanel.Controls.Add(this.InstructionsLabel);
            resources.ApplyResources(this.DescriptionPanel, "DescriptionPanel");
            this.DescriptionPanel.Name = "DescriptionPanel";
            // 
            // MembersGroupBox
            // 
            resources.ApplyResources(this.MembersGroupBox, "MembersGroupBox");
            this.MembersGroupBox.Controls.Add(this.InterfaceMembersGridView);
            this.MembersGroupBox.Controls.Add(this.DeselectAllButton);
            this.MembersGroupBox.Controls.Add(this.SelectAllButton);
            this.MembersGroupBox.Name = "MembersGroupBox";
            this.MembersGroupBox.TabStop = false;
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
            resources.ApplyResources(this.InterfaceMembersGridView, "InterfaceMembersGridView");
            this.InterfaceMembersGridView.MultiSelect = false;
            this.InterfaceMembersGridView.Name = "InterfaceMembersGridView";
            this.InterfaceMembersGridView.RowHeadersVisible = false;
            this.InterfaceMembersGridView.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing;
            this.InterfaceMembersGridView.RowTemplate.Height = 24;
            this.InterfaceMembersGridView.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.InterfaceMembersGridView.ShowEditingIcon = false;
            // 
            // DeselectAllButton
            // 
            resources.ApplyResources(this.DeselectAllButton, "DeselectAllButton");
            this.DeselectAllButton.Name = "DeselectAllButton";
            this.DeselectAllButton.UseVisualStyleBackColor = true;
            // 
            // SelectAllButton
            // 
            resources.ApplyResources(this.SelectAllButton, "SelectAllButton");
            this.SelectAllButton.Name = "SelectAllButton";
            this.SelectAllButton.UseVisualStyleBackColor = true;
            // 
            // ExtractInterfaceDialog
            // 
            this.AcceptButton = this.OkButton;
            resources.ApplyResources(this, "$this");
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.CancelButton = this.CancelDialogButton;
            this.Controls.Add(this.MembersGroupBox);
            this.Controls.Add(this.DescriptionPanel);
            this.Controls.Add(this.InvalidNameValidationIcon);
            this.Controls.Add(this.InterfaceNameBox);
            this.Controls.Add(this.NameLabel);
            this.Controls.Add(this.flowLayoutPanel2);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "ExtractInterfaceDialog";
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