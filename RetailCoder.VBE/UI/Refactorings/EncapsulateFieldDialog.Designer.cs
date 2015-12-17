namespace Rubberduck.UI.Refactorings
{
    partial class EncapsulateFieldDialog
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(EncapsulateFieldDialog));
            this.DescriptionPanel = new System.Windows.Forms.Panel();
            this.TitleLabel = new System.Windows.Forms.Label();
            this.InstructionsLabel = new System.Windows.Forms.Label();
            this.InvalidPropertyNameIcon = new System.Windows.Forms.PictureBox();
            this.PreviewBox = new System.Windows.Forms.TextBox();
            this.PreviewLabel = new System.Windows.Forms.Label();
            this.VariableAccessibilityComboBox = new System.Windows.Forms.ComboBox();
            this.AccessibilityLabel = new System.Windows.Forms.Label();
            this.PropertyNameTextBox = new System.Windows.Forms.TextBox();
            this.PropertyNameLabel = new System.Windows.Forms.Label();
            this.SetterTypeComboBox = new System.Windows.Forms.ComboBox();
            this.SetterTypeLabel = new System.Windows.Forms.Label();
            this.flowLayoutPanel2 = new System.Windows.Forms.FlowLayoutPanel();
            this.CancelDialogButton = new System.Windows.Forms.Button();
            this.OkButton = new System.Windows.Forms.Button();
            this.InvalidVariableNameIcon = new System.Windows.Forms.PictureBox();
            this.VariableNameTextBox = new System.Windows.Forms.TextBox();
            this.VariableNameLabel = new System.Windows.Forms.Label();
            this.DescriptionPanel.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.InvalidPropertyNameIcon)).BeginInit();
            this.flowLayoutPanel2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.InvalidVariableNameIcon)).BeginInit();
            this.SuspendLayout();
            // 
            // DescriptionPanel
            // 
            this.DescriptionPanel.BackColor = System.Drawing.Color.White;
            this.DescriptionPanel.Controls.Add(this.TitleLabel);
            this.DescriptionPanel.Controls.Add(this.InstructionsLabel);
            this.DescriptionPanel.Dock = System.Windows.Forms.DockStyle.Top;
            this.DescriptionPanel.Location = new System.Drawing.Point(0, 0);
            this.DescriptionPanel.Margin = new System.Windows.Forms.Padding(4);
            this.DescriptionPanel.Name = "DescriptionPanel";
            this.DescriptionPanel.Size = new System.Drawing.Size(734, 84);
            this.DescriptionPanel.TabIndex = 14;
            // 
            // TitleLabel
            // 
            this.TitleLabel.AutoSize = true;
            this.TitleLabel.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.TitleLabel.Location = new System.Drawing.Point(20, 11);
            this.TitleLabel.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.TitleLabel.Name = "TitleLabel";
            this.TitleLabel.Padding = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.TitleLabel.Size = new System.Drawing.Size(147, 22);
            this.TitleLabel.TabIndex = 2;
            this.TitleLabel.Text = "Encapsulate Field";
            // 
            // InstructionsLabel
            // 
            this.InstructionsLabel.Location = new System.Drawing.Point(20, 30);
            this.InstructionsLabel.Name = "InstructionsLabel";
            this.InstructionsLabel.Padding = new System.Windows.Forms.Padding(4);
            this.InstructionsLabel.Size = new System.Drawing.Size(549, 34);
            this.InstructionsLabel.TabIndex = 3;
            this.InstructionsLabel.Text = "Please specify property name, parameter accessibility, and setter type.";
            // 
            // InvalidPropertyNameIcon
            // 
            this.InvalidPropertyNameIcon.Image = global::Rubberduck.Properties.Resources.cross_circle;
            this.InvalidPropertyNameIcon.Location = new System.Drawing.Point(449, 104);
            this.InvalidPropertyNameIcon.Margin = new System.Windows.Forms.Padding(4);
            this.InvalidPropertyNameIcon.Name = "InvalidPropertyNameIcon";
            this.InvalidPropertyNameIcon.Size = new System.Drawing.Size(16, 16);
            this.InvalidPropertyNameIcon.SizeMode = System.Windows.Forms.PictureBoxSizeMode.AutoSize;
            this.InvalidPropertyNameIcon.TabIndex = 24;
            this.InvalidPropertyNameIcon.TabStop = false;
            // 
            // PreviewBox
            // 
            this.PreviewBox.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.PreviewBox.BackColor = System.Drawing.Color.White;
            this.PreviewBox.Font = new System.Drawing.Font("Consolas", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.PreviewBox.Location = new System.Drawing.Point(21, 221);
            this.PreviewBox.Margin = new System.Windows.Forms.Padding(4);
            this.PreviewBox.Multiline = true;
            this.PreviewBox.Name = "PreviewBox";
            this.PreviewBox.ReadOnly = true;
            this.PreviewBox.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.PreviewBox.Size = new System.Drawing.Size(691, 181);
            this.PreviewBox.TabIndex = 23;
            this.PreviewBox.WordWrap = false;
            // 
            // PreviewLabel
            // 
            this.PreviewLabel.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.PreviewLabel.AutoSize = true;
            this.PreviewLabel.Location = new System.Drawing.Point(20, 200);
            this.PreviewLabel.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.PreviewLabel.Name = "PreviewLabel";
            this.PreviewLabel.Size = new System.Drawing.Size(61, 17);
            this.PreviewLabel.TabIndex = 22;
            this.PreviewLabel.Text = "Preview:";
            // 
            // VariableAccessibilityComboBox
            // 
            this.VariableAccessibilityComboBox.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.VariableAccessibilityComboBox.FormattingEnabled = true;
            this.VariableAccessibilityComboBox.Location = new System.Drawing.Point(507, 164);
            this.VariableAccessibilityComboBox.Margin = new System.Windows.Forms.Padding(4);
            this.VariableAccessibilityComboBox.Name = "VariableAccessibilityComboBox";
            this.VariableAccessibilityComboBox.Size = new System.Drawing.Size(205, 24);
            this.VariableAccessibilityComboBox.TabIndex = 20;
            // 
            // AccessibilityLabel
            // 
            this.AccessibilityLabel.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.AccessibilityLabel.AutoSize = true;
            this.AccessibilityLabel.Location = new System.Drawing.Point(504, 143);
            this.AccessibilityLabel.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.AccessibilityLabel.Name = "AccessibilityLabel";
            this.AccessibilityLabel.Size = new System.Drawing.Size(88, 17);
            this.AccessibilityLabel.TabIndex = 19;
            this.AccessibilityLabel.Text = "Accessibility:";
            // 
            // PropertyNameTextBox
            // 
            this.PropertyNameTextBox.Location = new System.Drawing.Point(23, 113);
            this.PropertyNameTextBox.Margin = new System.Windows.Forms.Padding(4);
            this.PropertyNameTextBox.Name = "PropertyNameTextBox";
            this.PropertyNameTextBox.Size = new System.Drawing.Size(436, 22);
            this.PropertyNameTextBox.TabIndex = 16;
            // 
            // PropertyNameLabel
            // 
            this.PropertyNameLabel.AutoSize = true;
            this.PropertyNameLabel.Location = new System.Drawing.Point(18, 92);
            this.PropertyNameLabel.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.PropertyNameLabel.Name = "PropertyNameLabel";
            this.PropertyNameLabel.Size = new System.Drawing.Size(107, 17);
            this.PropertyNameLabel.TabIndex = 15;
            this.PropertyNameLabel.Text = "Property Name:";
            // 
            // SetterTypeComboBox
            // 
            this.SetterTypeComboBox.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.SetterTypeComboBox.FormattingEnabled = true;
            this.SetterTypeComboBox.Location = new System.Drawing.Point(507, 111);
            this.SetterTypeComboBox.Margin = new System.Windows.Forms.Padding(4);
            this.SetterTypeComboBox.Name = "SetterTypeComboBox";
            this.SetterTypeComboBox.Size = new System.Drawing.Size(205, 24);
            this.SetterTypeComboBox.TabIndex = 26;
            // 
            // SetterTypeLabel
            // 
            this.SetterTypeLabel.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.SetterTypeLabel.AutoSize = true;
            this.SetterTypeLabel.Location = new System.Drawing.Point(504, 90);
            this.SetterTypeLabel.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.SetterTypeLabel.Name = "SetterTypeLabel";
            this.SetterTypeLabel.Size = new System.Drawing.Size(86, 17);
            this.SetterTypeLabel.TabIndex = 25;
            this.SetterTypeLabel.Text = "Setter Type:";
            // 
            // flowLayoutPanel2
            // 
            this.flowLayoutPanel2.BackColor = System.Drawing.SystemColors.ControlDark;
            this.flowLayoutPanel2.Controls.Add(this.CancelDialogButton);
            this.flowLayoutPanel2.Controls.Add(this.OkButton);
            this.flowLayoutPanel2.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.flowLayoutPanel2.FlowDirection = System.Windows.Forms.FlowDirection.RightToLeft;
            this.flowLayoutPanel2.Location = new System.Drawing.Point(0, 410);
            this.flowLayoutPanel2.Margin = new System.Windows.Forms.Padding(4);
            this.flowLayoutPanel2.Name = "flowLayoutPanel2";
            this.flowLayoutPanel2.Padding = new System.Windows.Forms.Padding(11, 10, 0, 10);
            this.flowLayoutPanel2.Size = new System.Drawing.Size(734, 53);
            this.flowLayoutPanel2.TabIndex = 27;
            // 
            // CancelDialogButton
            // 
            this.CancelDialogButton.BackColor = System.Drawing.SystemColors.ButtonFace;
            this.CancelDialogButton.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.CancelDialogButton.Location = new System.Drawing.Point(619, 14);
            this.CancelDialogButton.Margin = new System.Windows.Forms.Padding(4);
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
            this.OkButton.Location = new System.Drawing.Point(511, 14);
            this.OkButton.Margin = new System.Windows.Forms.Padding(4);
            this.OkButton.Name = "OkButton";
            this.OkButton.Size = new System.Drawing.Size(100, 28);
            this.OkButton.TabIndex = 1;
            this.OkButton.Text = "Ok";
            this.OkButton.UseVisualStyleBackColor = false;
            // 
            // InvalidVariableNameIcon
            // 
            this.InvalidVariableNameIcon.Image = global::Rubberduck.Properties.Resources.cross_circle;
            this.InvalidVariableNameIcon.Location = new System.Drawing.Point(449, 157);
            this.InvalidVariableNameIcon.Margin = new System.Windows.Forms.Padding(4);
            this.InvalidVariableNameIcon.Name = "InvalidVariableNameIcon";
            this.InvalidVariableNameIcon.Size = new System.Drawing.Size(16, 16);
            this.InvalidVariableNameIcon.SizeMode = System.Windows.Forms.PictureBoxSizeMode.AutoSize;
            this.InvalidVariableNameIcon.TabIndex = 30;
            this.InvalidVariableNameIcon.TabStop = false;
            this.InvalidVariableNameIcon.Visible = false;
            // 
            // VariableNameTextBox
            // 
            this.VariableNameTextBox.Location = new System.Drawing.Point(23, 166);
            this.VariableNameTextBox.Margin = new System.Windows.Forms.Padding(4);
            this.VariableNameTextBox.Name = "VariableNameTextBox";
            this.VariableNameTextBox.Size = new System.Drawing.Size(436, 22);
            this.VariableNameTextBox.TabIndex = 29;
            this.VariableNameTextBox.Text = "value";
            // 
            // VariableNameLabel
            // 
            this.VariableNameLabel.AutoSize = true;
            this.VariableNameLabel.Location = new System.Drawing.Point(18, 143);
            this.VariableNameLabel.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.VariableNameLabel.Name = "VariableNameLabel";
            this.VariableNameLabel.Size = new System.Drawing.Size(105, 17);
            this.VariableNameLabel.TabIndex = 28;
            this.VariableNameLabel.Text = "Variable Name:";
            // 
            // EncapsulateFieldDialog
            // 
            this.AcceptButton = this.OkButton;
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.CancelButton = this.CancelDialogButton;
            this.ClientSize = new System.Drawing.Size(734, 463);
            this.Controls.Add(this.InvalidVariableNameIcon);
            this.Controls.Add(this.VariableNameTextBox);
            this.Controls.Add(this.VariableNameLabel);
            this.Controls.Add(this.flowLayoutPanel2);
            this.Controls.Add(this.SetterTypeComboBox);
            this.Controls.Add(this.SetterTypeLabel);
            this.Controls.Add(this.InvalidPropertyNameIcon);
            this.Controls.Add(this.PreviewBox);
            this.Controls.Add(this.PreviewLabel);
            this.Controls.Add(this.VariableAccessibilityComboBox);
            this.Controls.Add(this.AccessibilityLabel);
            this.Controls.Add(this.PropertyNameTextBox);
            this.Controls.Add(this.PropertyNameLabel);
            this.Controls.Add(this.DescriptionPanel);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "EncapsulateFieldDialog";
            this.Text = "Rubberduck - Encapsulate Field";
            this.DescriptionPanel.ResumeLayout(false);
            this.DescriptionPanel.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.InvalidPropertyNameIcon)).EndInit();
            this.flowLayoutPanel2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.InvalidVariableNameIcon)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Panel DescriptionPanel;
        private System.Windows.Forms.Label TitleLabel;
        private System.Windows.Forms.Label InstructionsLabel;
        private System.Windows.Forms.PictureBox InvalidPropertyNameIcon;
        private System.Windows.Forms.TextBox PreviewBox;
        private System.Windows.Forms.Label PreviewLabel;
        private System.Windows.Forms.ComboBox VariableAccessibilityComboBox;
        private System.Windows.Forms.Label AccessibilityLabel;
        private System.Windows.Forms.TextBox PropertyNameTextBox;
        private System.Windows.Forms.Label PropertyNameLabel;
        private System.Windows.Forms.ComboBox SetterTypeComboBox;
        private System.Windows.Forms.Label SetterTypeLabel;
        private System.Windows.Forms.FlowLayoutPanel flowLayoutPanel2;
        private System.Windows.Forms.Button CancelDialogButton;
        private System.Windows.Forms.Button OkButton;
        private System.Windows.Forms.PictureBox InvalidVariableNameIcon;
        private System.Windows.Forms.TextBox VariableNameTextBox;
        private System.Windows.Forms.Label VariableNameLabel;
    }
}