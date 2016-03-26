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
            this.PropertyNameTextBox = new System.Windows.Forms.TextBox();
            this.PropertyNameLabel = new System.Windows.Forms.Label();
            this.flowLayoutPanel2 = new System.Windows.Forms.FlowLayoutPanel();
            this.CancelDialogButton = new System.Windows.Forms.Button();
            this.OkButton = new System.Windows.Forms.Button();
            this.InvalidVariableNameIcon = new System.Windows.Forms.PictureBox();
            this.ParameterNameTextBox = new System.Windows.Forms.TextBox();
            this.VariableNameLabel = new System.Windows.Forms.Label();
            this.SetterTypeGroupBox = new System.Windows.Forms.GroupBox();
            this.SetSetterTypeCheckBox = new System.Windows.Forms.CheckBox();
            this.LetSetterTypeCheckBox = new System.Windows.Forms.CheckBox();
            this.DescriptionPanel.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.InvalidPropertyNameIcon)).BeginInit();
            this.flowLayoutPanel2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.InvalidVariableNameIcon)).BeginInit();
            this.SetterTypeGroupBox.SuspendLayout();
            this.SuspendLayout();
            // 
            // DescriptionPanel
            // 
            this.DescriptionPanel.BackColor = System.Drawing.Color.White;
            this.DescriptionPanel.Controls.Add(this.TitleLabel);
            this.DescriptionPanel.Controls.Add(this.InstructionsLabel);
            this.DescriptionPanel.Dock = System.Windows.Forms.DockStyle.Top;
            this.DescriptionPanel.Location = new System.Drawing.Point(0, 0);
            this.DescriptionPanel.Name = "DescriptionPanel";
            this.DescriptionPanel.Size = new System.Drawing.Size(500, 68);
            this.DescriptionPanel.TabIndex = 14;
            // 
            // TitleLabel
            // 
            this.TitleLabel.AutoSize = true;
            this.TitleLabel.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.TitleLabel.Location = new System.Drawing.Point(15, 9);
            this.TitleLabel.Name = "TitleLabel";
            this.TitleLabel.Padding = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.TitleLabel.Size = new System.Drawing.Size(126, 19);
            this.TitleLabel.TabIndex = 2;
            this.TitleLabel.Text = "Encapsulate Field";
            // 
            // InstructionsLabel
            // 
            this.InstructionsLabel.Location = new System.Drawing.Point(15, 24);
            this.InstructionsLabel.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.InstructionsLabel.Name = "InstructionsLabel";
            this.InstructionsLabel.Padding = new System.Windows.Forms.Padding(3, 3, 3, 3);
            this.InstructionsLabel.Size = new System.Drawing.Size(412, 28);
            this.InstructionsLabel.TabIndex = 3;
            this.InstructionsLabel.Text = "Please specify property name, parameter accessibility, and setter type.";
            // 
            // InvalidPropertyNameIcon
            // 
            this.InvalidPropertyNameIcon.Image = global::Rubberduck.Properties.Resources.cross_circle;
            this.InvalidPropertyNameIcon.Location = new System.Drawing.Point(355, 83);
            this.InvalidPropertyNameIcon.Name = "InvalidPropertyNameIcon";
            this.InvalidPropertyNameIcon.Size = new System.Drawing.Size(16, 16);
            this.InvalidPropertyNameIcon.SizeMode = System.Windows.Forms.PictureBoxSizeMode.AutoSize;
            this.InvalidPropertyNameIcon.TabIndex = 24;
            this.InvalidPropertyNameIcon.TabStop = false;
            // 
            // PreviewBox
            // 
            this.PreviewBox.BackColor = System.Drawing.Color.White;
            this.PreviewBox.Font = new System.Drawing.Font("Consolas", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.PreviewBox.Location = new System.Drawing.Point(16, 180);
            this.PreviewBox.Multiline = true;
            this.PreviewBox.Name = "PreviewBox";
            this.PreviewBox.ReadOnly = true;
            this.PreviewBox.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.PreviewBox.Size = new System.Drawing.Size(467, 177);
            this.PreviewBox.TabIndex = 23;
            this.PreviewBox.WordWrap = false;
            // 
            // PreviewLabel
            // 
            this.PreviewLabel.AutoSize = true;
            this.PreviewLabel.Location = new System.Drawing.Point(15, 162);
            this.PreviewLabel.Name = "PreviewLabel";
            this.PreviewLabel.Size = new System.Drawing.Size(48, 13);
            this.PreviewLabel.TabIndex = 22;
            this.PreviewLabel.Text = "Preview:";
            // 
            // PropertyNameTextBox
            // 
            this.PropertyNameTextBox.Location = new System.Drawing.Point(17, 92);
            this.PropertyNameTextBox.Name = "PropertyNameTextBox";
            this.PropertyNameTextBox.Size = new System.Drawing.Size(344, 20);
            this.PropertyNameTextBox.TabIndex = 16;
            // 
            // PropertyNameLabel
            // 
            this.PropertyNameLabel.AutoSize = true;
            this.PropertyNameLabel.Location = new System.Drawing.Point(14, 75);
            this.PropertyNameLabel.Name = "PropertyNameLabel";
            this.PropertyNameLabel.Size = new System.Drawing.Size(80, 13);
            this.PropertyNameLabel.TabIndex = 15;
            this.PropertyNameLabel.Text = "Property Name:";
            // 
            // flowLayoutPanel2
            // 
            this.flowLayoutPanel2.BackColor = System.Drawing.SystemColors.ControlDark;
            this.flowLayoutPanel2.Controls.Add(this.CancelDialogButton);
            this.flowLayoutPanel2.Controls.Add(this.OkButton);
            this.flowLayoutPanel2.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.flowLayoutPanel2.FlowDirection = System.Windows.Forms.FlowDirection.RightToLeft;
            this.flowLayoutPanel2.Location = new System.Drawing.Point(0, 362);
            this.flowLayoutPanel2.Name = "flowLayoutPanel2";
            this.flowLayoutPanel2.Padding = new System.Windows.Forms.Padding(8, 8, 0, 8);
            this.flowLayoutPanel2.Size = new System.Drawing.Size(500, 43);
            this.flowLayoutPanel2.TabIndex = 27;
            // 
            // CancelDialogButton
            // 
            this.CancelDialogButton.BackColor = System.Drawing.SystemColors.ButtonFace;
            this.CancelDialogButton.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.CancelDialogButton.Location = new System.Drawing.Point(414, 11);
            this.CancelDialogButton.Name = "CancelDialogButton";
            this.CancelDialogButton.Size = new System.Drawing.Size(75, 23);
            this.CancelDialogButton.TabIndex = 0;
            this.CancelDialogButton.Text = "Cancel";
            this.CancelDialogButton.UseVisualStyleBackColor = false;
            // 
            // OkButton
            // 
            this.OkButton.BackColor = System.Drawing.SystemColors.ButtonFace;
            this.OkButton.DialogResult = System.Windows.Forms.DialogResult.OK;
            this.OkButton.Location = new System.Drawing.Point(333, 11);
            this.OkButton.Name = "OkButton";
            this.OkButton.Size = new System.Drawing.Size(75, 23);
            this.OkButton.TabIndex = 1;
            this.OkButton.Text = "Ok";
            this.OkButton.UseVisualStyleBackColor = false;
            // 
            // InvalidVariableNameIcon
            // 
            this.InvalidVariableNameIcon.Image = global::Rubberduck.Properties.Resources.cross_circle;
            this.InvalidVariableNameIcon.Location = new System.Drawing.Point(355, 127);
            this.InvalidVariableNameIcon.Name = "InvalidVariableNameIcon";
            this.InvalidVariableNameIcon.Size = new System.Drawing.Size(16, 16);
            this.InvalidVariableNameIcon.SizeMode = System.Windows.Forms.PictureBoxSizeMode.AutoSize;
            this.InvalidVariableNameIcon.TabIndex = 30;
            this.InvalidVariableNameIcon.TabStop = false;
            this.InvalidVariableNameIcon.Visible = false;
            // 
            // ParameterNameTextBox
            // 
            this.ParameterNameTextBox.Location = new System.Drawing.Point(17, 135);
            this.ParameterNameTextBox.Name = "ParameterNameTextBox";
            this.ParameterNameTextBox.Size = new System.Drawing.Size(344, 20);
            this.ParameterNameTextBox.TabIndex = 29;
            this.ParameterNameTextBox.Text = "value";
            // 
            // VariableNameLabel
            // 
            this.VariableNameLabel.AutoSize = true;
            this.VariableNameLabel.Location = new System.Drawing.Point(14, 116);
            this.VariableNameLabel.Name = "VariableNameLabel";
            this.VariableNameLabel.Size = new System.Drawing.Size(89, 13);
            this.VariableNameLabel.TabIndex = 28;
            this.VariableNameLabel.Text = "Parameter Name:";
            // 
            // SetterTypeGroupBox
            // 
            this.SetterTypeGroupBox.Controls.Add(this.SetSetterTypeCheckBox);
            this.SetterTypeGroupBox.Controls.Add(this.LetSetterTypeCheckBox);
            this.SetterTypeGroupBox.Location = new System.Drawing.Point(376, 84);
            this.SetterTypeGroupBox.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.SetterTypeGroupBox.Name = "SetterTypeGroupBox";
            this.SetterTypeGroupBox.Padding = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.SetterTypeGroupBox.Size = new System.Drawing.Size(106, 70);
            this.SetterTypeGroupBox.TabIndex = 31;
            this.SetterTypeGroupBox.TabStop = false;
            this.SetterTypeGroupBox.Text = "Assignment:";
            // 
            // SetSetterTypeCheckBox
            // 
            this.SetSetterTypeCheckBox.AutoSize = true;
            this.SetSetterTypeCheckBox.Location = new System.Drawing.Point(10, 39);
            this.SetSetterTypeCheckBox.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.SetSetterTypeCheckBox.Name = "SetSetterTypeCheckBox";
            this.SetSetterTypeCheckBox.Size = new System.Drawing.Size(42, 17);
            this.SetSetterTypeCheckBox.TabIndex = 3;
            this.SetSetterTypeCheckBox.Text = "Set";
            this.SetSetterTypeCheckBox.UseVisualStyleBackColor = true;
            // 
            // LetSetterTypeCheckBox
            // 
            this.LetSetterTypeCheckBox.AutoSize = true;
            this.LetSetterTypeCheckBox.Location = new System.Drawing.Point(10, 17);
            this.LetSetterTypeCheckBox.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.LetSetterTypeCheckBox.Name = "LetSetterTypeCheckBox";
            this.LetSetterTypeCheckBox.Size = new System.Drawing.Size(41, 17);
            this.LetSetterTypeCheckBox.TabIndex = 2;
            this.LetSetterTypeCheckBox.Text = "Let";
            this.LetSetterTypeCheckBox.UseVisualStyleBackColor = true;
            // 
            // EncapsulateFieldDialog
            // 
            this.AcceptButton = this.OkButton;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.CancelButton = this.CancelDialogButton;
            this.ClientSize = new System.Drawing.Size(500, 405);
            this.Controls.Add(this.SetterTypeGroupBox);
            this.Controls.Add(this.InvalidVariableNameIcon);
            this.Controls.Add(this.ParameterNameTextBox);
            this.Controls.Add(this.VariableNameLabel);
            this.Controls.Add(this.flowLayoutPanel2);
            this.Controls.Add(this.InvalidPropertyNameIcon);
            this.Controls.Add(this.PreviewBox);
            this.Controls.Add(this.PreviewLabel);
            this.Controls.Add(this.PropertyNameTextBox);
            this.Controls.Add(this.PropertyNameLabel);
            this.Controls.Add(this.DescriptionPanel);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "EncapsulateFieldDialog";
            this.ShowInTaskbar = false;
            this.Text = "Rubberduck - Encapsulate Field";
            this.DescriptionPanel.ResumeLayout(false);
            this.DescriptionPanel.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.InvalidPropertyNameIcon)).EndInit();
            this.flowLayoutPanel2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.InvalidVariableNameIcon)).EndInit();
            this.SetterTypeGroupBox.ResumeLayout(false);
            this.SetterTypeGroupBox.PerformLayout();
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
        private System.Windows.Forms.TextBox PropertyNameTextBox;
        private System.Windows.Forms.Label PropertyNameLabel;
        private System.Windows.Forms.FlowLayoutPanel flowLayoutPanel2;
        private System.Windows.Forms.Button CancelDialogButton;
        private System.Windows.Forms.Button OkButton;
        private System.Windows.Forms.PictureBox InvalidVariableNameIcon;
        private System.Windows.Forms.TextBox ParameterNameTextBox;
        private System.Windows.Forms.Label VariableNameLabel;
        private System.Windows.Forms.GroupBox SetterTypeGroupBox;
        private System.Windows.Forms.CheckBox SetSetterTypeCheckBox;
        private System.Windows.Forms.CheckBox LetSetterTypeCheckBox;
    }
}