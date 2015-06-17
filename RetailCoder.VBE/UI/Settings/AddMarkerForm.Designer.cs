namespace Rubberduck.UI.Settings
{
    partial class AddMarkerForm
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(AddMarkerForm));
            this.TodoMarkerPriorityComboBox = new System.Windows.Forms.ComboBox();
            this.TodoMarkerTextBox = new System.Windows.Forms.TextBox();
            this.TodoMarkerTextBoxLabel = new System.Windows.Forms.Label();
            this.TodoMarkerPriorityComboBoxLabel = new System.Windows.Forms.Label();
            this.flowLayoutPanel2 = new System.Windows.Forms.FlowLayoutPanel();
            this.CancelButton = new System.Windows.Forms.Button();
            this.OkButton = new System.Windows.Forms.Button();
            this.InvalidNameValidationIcon = new System.Windows.Forms.PictureBox();
            this.flowLayoutPanel2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.InvalidNameValidationIcon)).BeginInit();
            this.SuspendLayout();
            // 
            // TodoMarkerPriorityComboBox
            // 
            this.TodoMarkerPriorityComboBox.FormattingEnabled = true;
            this.TodoMarkerPriorityComboBox.Location = new System.Drawing.Point(12, 88);
            this.TodoMarkerPriorityComboBox.Name = "TodoMarkerPriorityComboBox";
            this.TodoMarkerPriorityComboBox.Size = new System.Drawing.Size(353, 24);
            this.TodoMarkerPriorityComboBox.TabIndex = 0;
            // 
            // TodoMarkerTextBox
            // 
            this.TodoMarkerTextBox.Location = new System.Drawing.Point(12, 29);
            this.TodoMarkerTextBox.Name = "TodoMarkerTextBox";
            this.TodoMarkerTextBox.Size = new System.Drawing.Size(353, 22);
            this.TodoMarkerTextBox.TabIndex = 1;
            // 
            // TodoMarkerTextBoxLabel
            // 
            this.TodoMarkerTextBoxLabel.AutoSize = true;
            this.TodoMarkerTextBoxLabel.Location = new System.Drawing.Point(12, 9);
            this.TodoMarkerTextBoxLabel.Name = "TodoMarkerTextBoxLabel";
            this.TodoMarkerTextBoxLabel.Size = new System.Drawing.Size(39, 17);
            this.TodoMarkerTextBoxLabel.TabIndex = 2;
            this.TodoMarkerTextBoxLabel.Text = "Text:";
            // 
            // TodoMarkerPriorityComboBoxLabel
            // 
            this.TodoMarkerPriorityComboBoxLabel.AutoSize = true;
            this.TodoMarkerPriorityComboBoxLabel.Location = new System.Drawing.Point(12, 68);
            this.TodoMarkerPriorityComboBoxLabel.Name = "TodoMarkerPriorityComboBoxLabel";
            this.TodoMarkerPriorityComboBoxLabel.Size = new System.Drawing.Size(56, 17);
            this.TodoMarkerPriorityComboBoxLabel.TabIndex = 3;
            this.TodoMarkerPriorityComboBoxLabel.Text = "Priority:";
            // 
            // flowLayoutPanel2
            // 
            this.flowLayoutPanel2.BackColor = System.Drawing.SystemColors.ControlDark;
            this.flowLayoutPanel2.Controls.Add(this.CancelButton);
            this.flowLayoutPanel2.Controls.Add(this.OkButton);
            this.flowLayoutPanel2.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.flowLayoutPanel2.FlowDirection = System.Windows.Forms.FlowDirection.RightToLeft;
            this.flowLayoutPanel2.Location = new System.Drawing.Point(0, 150);
            this.flowLayoutPanel2.Margin = new System.Windows.Forms.Padding(4);
            this.flowLayoutPanel2.Name = "flowLayoutPanel2";
            this.flowLayoutPanel2.Padding = new System.Windows.Forms.Padding(11, 10, 0, 10);
            this.flowLayoutPanel2.Size = new System.Drawing.Size(380, 53);
            this.flowLayoutPanel2.TabIndex = 4;
            // 
            // CancelButton
            // 
            this.CancelButton.BackColor = System.Drawing.SystemColors.ButtonFace;
            this.CancelButton.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.CancelButton.Location = new System.Drawing.Point(265, 14);
            this.CancelButton.Margin = new System.Windows.Forms.Padding(4);
            this.CancelButton.Name = "CancelButton";
            this.CancelButton.Size = new System.Drawing.Size(100, 28);
            this.CancelButton.TabIndex = 0;
            this.CancelButton.Text = "Cancel";
            this.CancelButton.UseVisualStyleBackColor = false;
            // 
            // OkButton
            // 
            this.OkButton.BackColor = System.Drawing.SystemColors.ButtonFace;
            this.OkButton.DialogResult = System.Windows.Forms.DialogResult.OK;
            this.OkButton.Location = new System.Drawing.Point(157, 14);
            this.OkButton.Margin = new System.Windows.Forms.Padding(4);
            this.OkButton.Name = "OkButton";
            this.OkButton.Size = new System.Drawing.Size(100, 28);
            this.OkButton.TabIndex = 1;
            this.OkButton.Text = "Ok";
            this.OkButton.UseVisualStyleBackColor = false;
            // 
            // InvalidNameValidationIcon
            // 
            this.InvalidNameValidationIcon.Image = global::Rubberduck.Properties.Resources.cross_circle;
            this.InvalidNameValidationIcon.Location = new System.Drawing.Point(356, 22);
            this.InvalidNameValidationIcon.Margin = new System.Windows.Forms.Padding(4);
            this.InvalidNameValidationIcon.Name = "InvalidNameValidationIcon";
            this.InvalidNameValidationIcon.Size = new System.Drawing.Size(16, 16);
            this.InvalidNameValidationIcon.SizeMode = System.Windows.Forms.PictureBoxSizeMode.AutoSize;
            this.InvalidNameValidationIcon.TabIndex = 14;
            this.InvalidNameValidationIcon.TabStop = false;
            // 
            // AddMarkerForm
            // 
            this.AcceptButton = this.OkButton;
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.CancelButton = this.CancelButton;
            this.ClientSize = new System.Drawing.Size(380, 203);
            this.Controls.Add(this.InvalidNameValidationIcon);
            this.Controls.Add(this.flowLayoutPanel2);
            this.Controls.Add(this.TodoMarkerPriorityComboBoxLabel);
            this.Controls.Add(this.TodoMarkerTextBoxLabel);
            this.Controls.Add(this.TodoMarkerTextBox);
            this.Controls.Add(this.TodoMarkerPriorityComboBox);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "AddMarkerForm";
            this.ShowInTaskbar = false;
            this.Text = "AddMarkerForm";
            this.TopMost = true;
            this.flowLayoutPanel2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.InvalidNameValidationIcon)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.ComboBox TodoMarkerPriorityComboBox;
        private System.Windows.Forms.TextBox TodoMarkerTextBox;
        private System.Windows.Forms.Label TodoMarkerTextBoxLabel;
        private System.Windows.Forms.Label TodoMarkerPriorityComboBoxLabel;
        private System.Windows.Forms.FlowLayoutPanel flowLayoutPanel2;
        private System.Windows.Forms.Button CancelButton;
        private System.Windows.Forms.Button OkButton;
        private System.Windows.Forms.PictureBox InvalidNameValidationIcon;
    }
}