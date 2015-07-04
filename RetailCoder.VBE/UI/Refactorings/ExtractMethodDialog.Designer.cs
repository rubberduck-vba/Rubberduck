using System.ComponentModel;
using System.Windows.Forms;

namespace Rubberduck.UI.Refactorings
{
    partial class ExtractMethodDialog
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

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.flowLayoutPanel2 = new System.Windows.Forms.FlowLayoutPanel();
            this.CancelDialogButton = new System.Windows.Forms.Button();
            this.OkButton = new System.Windows.Forms.Button();
            this.panel2 = new System.Windows.Forms.Panel();
            this.TitleLabel = new System.Windows.Forms.Label();
            this.InstructionsLabel = new System.Windows.Forms.Label();
            this.panel1 = new System.Windows.Forms.Panel();
            this.SetReturnValueCheck = new System.Windows.Forms.CheckBox();
            this.InvalidNameValidationIcon = new System.Windows.Forms.PictureBox();
            this.PreviewBox = new System.Windows.Forms.TextBox();
            this.PreviewLabel = new System.Windows.Forms.Label();
            this.MethodParametersGrid = new System.Windows.Forms.DataGridView();
            this.ParametersLabel = new System.Windows.Forms.Label();
            this.MethodAccessibilityCombo = new System.Windows.Forms.ComboBox();
            this.AccessibilityLabel = new System.Windows.Forms.Label();
            this.MethodReturnValueCombo = new System.Windows.Forms.ComboBox();
            this.ReturnLabel = new System.Windows.Forms.Label();
            this.MethodNameBox = new System.Windows.Forms.TextBox();
            this.NameLabel = new System.Windows.Forms.Label();
            this.flowLayoutPanel2.SuspendLayout();
            this.panel2.SuspendLayout();
            this.panel1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.InvalidNameValidationIcon)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.MethodParametersGrid)).BeginInit();
            this.SuspendLayout();
            // 
            // flowLayoutPanel2
            // 
            this.flowLayoutPanel2.BackColor = System.Drawing.SystemColors.ControlDark;
            this.flowLayoutPanel2.Controls.Add(this.CancelDialogButton);
            this.flowLayoutPanel2.Controls.Add(this.OkButton);
            this.flowLayoutPanel2.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.flowLayoutPanel2.FlowDirection = System.Windows.Forms.FlowDirection.RightToLeft;
            this.flowLayoutPanel2.Location = new System.Drawing.Point(0, 539);
            this.flowLayoutPanel2.Margin = new System.Windows.Forms.Padding(4);
            this.flowLayoutPanel2.Name = "flowLayoutPanel2";
            this.flowLayoutPanel2.Padding = new System.Windows.Forms.Padding(11, 10, 0, 10);
            this.flowLayoutPanel2.Size = new System.Drawing.Size(768, 53);
            this.flowLayoutPanel2.TabIndex = 1;
            // 
            // CancelDialogButton
            // 
            this.CancelDialogButton.BackColor = System.Drawing.SystemColors.ButtonFace;
            this.CancelDialogButton.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.CancelDialogButton.Location = new System.Drawing.Point(653, 14);
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
            this.OkButton.Location = new System.Drawing.Point(545, 14);
            this.OkButton.Margin = new System.Windows.Forms.Padding(4);
            this.OkButton.Name = "OkButton";
            this.OkButton.Size = new System.Drawing.Size(100, 28);
            this.OkButton.TabIndex = 1;
            this.OkButton.Text = "Ok";
            this.OkButton.UseVisualStyleBackColor = false;
            // 
            // panel2
            // 
            this.panel2.BackColor = System.Drawing.Color.White;
            this.panel2.Controls.Add(this.TitleLabel);
            this.panel2.Controls.Add(this.InstructionsLabel);
            this.panel2.Dock = System.Windows.Forms.DockStyle.Top;
            this.panel2.Location = new System.Drawing.Point(0, 0);
            this.panel2.Margin = new System.Windows.Forms.Padding(4);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(768, 84);
            this.panel2.TabIndex = 13;
            // 
            // TitleLabel
            // 
            this.TitleLabel.AutoSize = true;
            this.TitleLabel.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.TitleLabel.Location = new System.Drawing.Point(20, 11);
            this.TitleLabel.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.TitleLabel.Name = "TitleLabel";
            this.TitleLabel.Padding = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.TitleLabel.Size = new System.Drawing.Size(128, 22);
            this.TitleLabel.TabIndex = 2;
            this.TitleLabel.Text = "Extract Method";
            // 
            // InstructionsLabel
            // 
            this.InstructionsLabel.AutoSize = true;
            this.InstructionsLabel.Location = new System.Drawing.Point(16, 37);
            this.InstructionsLabel.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.InstructionsLabel.Name = "InstructionsLabel";
            this.InstructionsLabel.Padding = new System.Windows.Forms.Padding(5);
            this.InstructionsLabel.Size = new System.Drawing.Size(609, 27);
            this.InstructionsLabel.TabIndex = 3;
            this.InstructionsLabel.Text = "Please specify method name, return type and/or parameters (if applicable), and ot" +
    "her options.";
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.SetReturnValueCheck);
            this.panel1.Controls.Add(this.InvalidNameValidationIcon);
            this.panel1.Controls.Add(this.PreviewBox);
            this.panel1.Controls.Add(this.PreviewLabel);
            this.panel1.Controls.Add(this.MethodParametersGrid);
            this.panel1.Controls.Add(this.ParametersLabel);
            this.panel1.Controls.Add(this.MethodAccessibilityCombo);
            this.panel1.Controls.Add(this.AccessibilityLabel);
            this.panel1.Controls.Add(this.MethodReturnValueCombo);
            this.panel1.Controls.Add(this.ReturnLabel);
            this.panel1.Controls.Add(this.MethodNameBox);
            this.panel1.Controls.Add(this.NameLabel);
            this.panel1.Location = new System.Drawing.Point(0, 87);
            this.panel1.Margin = new System.Windows.Forms.Padding(4);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(768, 458);
            this.panel1.TabIndex = 14;
            // 
            // SetReturnValueCheck
            // 
            this.SetReturnValueCheck.AutoSize = true;
            this.SetReturnValueCheck.Location = new System.Drawing.Point(340, 46);
            this.SetReturnValueCheck.Margin = new System.Windows.Forms.Padding(4);
            this.SetReturnValueCheck.Name = "SetReturnValueCheck";
            this.SetReturnValueCheck.Size = new System.Drawing.Size(51, 21);
            this.SetReturnValueCheck.TabIndex = 11;
            this.SetReturnValueCheck.Text = "Set";
            this.SetReturnValueCheck.UseVisualStyleBackColor = true;
            // 
            // InvalidNameValidationIcon
            // 
            this.InvalidNameValidationIcon.Image = global::Rubberduck.Properties.Resources.cross_circle;
            this.InvalidNameValidationIcon.Location = new System.Drawing.Point(743, 4);
            this.InvalidNameValidationIcon.Margin = new System.Windows.Forms.Padding(4);
            this.InvalidNameValidationIcon.Name = "InvalidNameValidationIcon";
            this.InvalidNameValidationIcon.Size = new System.Drawing.Size(16, 16);
            this.InvalidNameValidationIcon.SizeMode = System.Windows.Forms.PictureBoxSizeMode.AutoSize;
            this.InvalidNameValidationIcon.TabIndex = 10;
            this.InvalidNameValidationIcon.TabStop = false;
            // 
            // PreviewBox
            // 
            this.PreviewBox.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.PreviewBox.BackColor = System.Drawing.Color.White;
            this.PreviewBox.Font = new System.Drawing.Font("Consolas", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.PreviewBox.Location = new System.Drawing.Point(24, 254);
            this.PreviewBox.Margin = new System.Windows.Forms.Padding(4);
            this.PreviewBox.Multiline = true;
            this.PreviewBox.Name = "PreviewBox";
            this.PreviewBox.ReadOnly = true;
            this.PreviewBox.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.PreviewBox.Size = new System.Drawing.Size(727, 181);
            this.PreviewBox.TabIndex = 9;
            this.PreviewBox.WordWrap = false;
            // 
            // PreviewLabel
            // 
            this.PreviewLabel.AutoSize = true;
            this.PreviewLabel.Location = new System.Drawing.Point(20, 234);
            this.PreviewLabel.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.PreviewLabel.Name = "PreviewLabel";
            this.PreviewLabel.Size = new System.Drawing.Size(61, 17);
            this.PreviewLabel.TabIndex = 8;
            this.PreviewLabel.Text = "Preview:";
            // 
            // MethodParametersGrid
            // 
            this.MethodParametersGrid.AllowUserToAddRows = false;
            this.MethodParametersGrid.AllowUserToDeleteRows = false;
            this.MethodParametersGrid.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right)));
            this.MethodParametersGrid.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.MethodParametersGrid.Location = new System.Drawing.Point(24, 101);
            this.MethodParametersGrid.Margin = new System.Windows.Forms.Padding(11, 4, 11, 4);
            this.MethodParametersGrid.Name = "MethodParametersGrid";
            this.MethodParametersGrid.Size = new System.Drawing.Size(728, 119);
            this.MethodParametersGrid.TabIndex = 7;
            // 
            // ParametersLabel
            // 
            this.ParametersLabel.AutoSize = true;
            this.ParametersLabel.Location = new System.Drawing.Point(20, 81);
            this.ParametersLabel.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.ParametersLabel.Name = "ParametersLabel";
            this.ParametersLabel.Size = new System.Drawing.Size(85, 17);
            this.ParametersLabel.TabIndex = 6;
            this.ParametersLabel.Text = "Parameters:";
            // 
            // MethodAccessibilityCombo
            // 
            this.MethodAccessibilityCombo.FormattingEnabled = true;
            this.MethodAccessibilityCombo.Location = new System.Drawing.Point(547, 42);
            this.MethodAccessibilityCombo.Margin = new System.Windows.Forms.Padding(4);
            this.MethodAccessibilityCombo.Name = "MethodAccessibilityCombo";
            this.MethodAccessibilityCombo.Size = new System.Drawing.Size(205, 24);
            this.MethodAccessibilityCombo.TabIndex = 5;
            // 
            // AccessibilityLabel
            // 
            this.AccessibilityLabel.AutoSize = true;
            this.AccessibilityLabel.Location = new System.Drawing.Point(448, 46);
            this.AccessibilityLabel.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.AccessibilityLabel.Name = "AccessibilityLabel";
            this.AccessibilityLabel.Size = new System.Drawing.Size(88, 17);
            this.AccessibilityLabel.TabIndex = 4;
            this.AccessibilityLabel.Text = "Accessibility:";
            // 
            // MethodReturnValueCombo
            // 
            this.MethodReturnValueCombo.FormattingEnabled = true;
            this.MethodReturnValueCombo.Location = new System.Drawing.Point(84, 42);
            this.MethodReturnValueCombo.Margin = new System.Windows.Forms.Padding(4);
            this.MethodReturnValueCombo.Name = "MethodReturnValueCombo";
            this.MethodReturnValueCombo.Size = new System.Drawing.Size(245, 24);
            this.MethodReturnValueCombo.TabIndex = 3;
            // 
            // ReturnLabel
            // 
            this.ReturnLabel.AutoSize = true;
            this.ReturnLabel.Location = new System.Drawing.Point(20, 46);
            this.ReturnLabel.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.ReturnLabel.Name = "ReturnLabel";
            this.ReturnLabel.Size = new System.Drawing.Size(55, 17);
            this.ReturnLabel.TabIndex = 2;
            this.ReturnLabel.Text = "Return:";
            // 
            // MethodNameBox
            // 
            this.MethodNameBox.Location = new System.Drawing.Point(84, 9);
            this.MethodNameBox.Margin = new System.Windows.Forms.Padding(4);
            this.MethodNameBox.Name = "MethodNameBox";
            this.MethodNameBox.Size = new System.Drawing.Size(667, 22);
            this.MethodNameBox.TabIndex = 1;
            // 
            // NameLabel
            // 
            this.NameLabel.AutoSize = true;
            this.NameLabel.Location = new System.Drawing.Point(20, 12);
            this.NameLabel.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.NameLabel.Name = "NameLabel";
            this.NameLabel.Size = new System.Drawing.Size(49, 17);
            this.NameLabel.TabIndex = 0;
            this.NameLabel.Text = "Name:";
            // 
            // ExtractMethodDialog
            // 
            this.AcceptButton = this.OkButton;
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.CancelButton = this.CancelDialogButton;
            this.ClientSize = new System.Drawing.Size(768, 592);
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.panel2);
            this.Controls.Add(this.flowLayoutPanel2);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.Icon = global::Rubberduck.UI.RubberduckUI.Ducky;
            this.Margin = new System.Windows.Forms.Padding(4);
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "ExtractMethodDialog";
            this.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide;
            this.Text = "Rubberduck - Extract Method";
            this.flowLayoutPanel2.ResumeLayout(false);
            this.panel2.ResumeLayout(false);
            this.panel2.PerformLayout();
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.InvalidNameValidationIcon)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.MethodParametersGrid)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private FlowLayoutPanel flowLayoutPanel2;
        private Button CancelDialogButton;
        private Button OkButton;
        private Panel panel2;
        private Label TitleLabel;
        private Label InstructionsLabel;
        private Panel panel1;
        private CheckBox SetReturnValueCheck;
        private PictureBox InvalidNameValidationIcon;
        private TextBox PreviewBox;
        private Label PreviewLabel;
        private DataGridView MethodParametersGrid;
        private Label ParametersLabel;
        private ComboBox MethodAccessibilityCombo;
        private Label AccessibilityLabel;
        private ComboBox MethodReturnValueCombo;
        private Label ReturnLabel;
        private TextBox MethodNameBox;
        private Label NameLabel;
    }
}