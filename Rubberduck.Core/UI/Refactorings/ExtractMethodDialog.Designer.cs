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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(ExtractMethodDialog));
            this.flowLayoutPanel2 = new System.Windows.Forms.FlowLayoutPanel();
            this.CancelDialogButton = new System.Windows.Forms.Button();
            this.OkButton = new System.Windows.Forms.Button();
            this.panel2 = new System.Windows.Forms.Panel();
            this.TitleLabel = new System.Windows.Forms.Label();
            this.InstructionsLabel = new System.Windows.Forms.Label();
            this.panel1 = new System.Windows.Forms.Panel();
            this.InvalidNameValidationIcon = new System.Windows.Forms.PictureBox();
            this.PreviewBox = new System.Windows.Forms.TextBox();
            this.PreviewLabel = new System.Windows.Forms.Label();
            this.MethodParametersGrid = new System.Windows.Forms.DataGridView();
            this.ParametersLabel = new System.Windows.Forms.Label();
            this.MethodAccessibilityCombo = new System.Windows.Forms.ComboBox();
            this.AccessibilityLabel = new System.Windows.Forms.Label();
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
            this.flowLayoutPanel2.Location = new System.Drawing.Point(0, 438);
            this.flowLayoutPanel2.Name = "flowLayoutPanel2";
            this.flowLayoutPanel2.Padding = new System.Windows.Forms.Padding(8, 8, 0, 8);
            this.flowLayoutPanel2.Size = new System.Drawing.Size(600, 43);
            this.flowLayoutPanel2.TabIndex = 1;
            // 
            // CancelDialogButton
            // 
            this.CancelDialogButton.BackColor = System.Drawing.SystemColors.ButtonFace;
            this.CancelDialogButton.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.CancelDialogButton.Location = new System.Drawing.Point(514, 11);
            this.CancelDialogButton.Name = "CancelDialogButton";
            this.CancelDialogButton.Size = new System.Drawing.Size(75, 23);
            this.CancelDialogButton.TabIndex = 4;
            this.CancelDialogButton.Text = "Cancel";
            this.CancelDialogButton.UseVisualStyleBackColor = false;
            // 
            // OkButton
            // 
            this.OkButton.BackColor = System.Drawing.SystemColors.ButtonFace;
            this.OkButton.DialogResult = System.Windows.Forms.DialogResult.OK;
            this.OkButton.Location = new System.Drawing.Point(433, 11);
            this.OkButton.Name = "OkButton";
            this.OkButton.Size = new System.Drawing.Size(75, 23);
            this.OkButton.TabIndex = 3;
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
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(600, 68);
            this.panel2.TabIndex = 13;
            // 
            // TitleLabel
            // 
            this.TitleLabel.AutoSize = true;
            this.TitleLabel.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.TitleLabel.Location = new System.Drawing.Point(15, 9);
            this.TitleLabel.Name = "TitleLabel";
            this.TitleLabel.Padding = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.TitleLabel.Size = new System.Drawing.Size(107, 19);
            this.TitleLabel.TabIndex = 2;
            this.TitleLabel.Text = "Extract Method";
            // 
            // InstructionsLabel
            // 
            this.InstructionsLabel.Location = new System.Drawing.Point(15, 24);
            this.InstructionsLabel.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.InstructionsLabel.Name = "InstructionsLabel";
            this.InstructionsLabel.Padding = new System.Windows.Forms.Padding(3, 3, 3, 3);
            this.InstructionsLabel.Size = new System.Drawing.Size(412, 28);
            this.InstructionsLabel.TabIndex = 3;
            this.InstructionsLabel.Text = "Please specify method name, return type and/or parameters (if applicable), and ot" +
    "her options.";
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.InvalidNameValidationIcon);
            this.panel1.Controls.Add(this.PreviewBox);
            this.panel1.Controls.Add(this.PreviewLabel);
            this.panel1.Controls.Add(this.MethodParametersGrid);
            this.panel1.Controls.Add(this.ParametersLabel);
            this.panel1.Controls.Add(this.MethodAccessibilityCombo);
            this.panel1.Controls.Add(this.AccessibilityLabel);
            this.panel1.Controls.Add(this.MethodNameBox);
            this.panel1.Controls.Add(this.NameLabel);
            this.panel1.Location = new System.Drawing.Point(0, 71);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(600, 372);
            this.panel1.TabIndex = 14;
            // 
            // InvalidNameValidationIcon
            // 
            this.InvalidNameValidationIcon.Image = global::Rubberduck.Properties.Resources.cross_circle;
            this.InvalidNameValidationIcon.Location = new System.Drawing.Point(577, 1);
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
            this.PreviewBox.Location = new System.Drawing.Point(18, 206);
            this.PreviewBox.Multiline = true;
            this.PreviewBox.Name = "PreviewBox";
            this.PreviewBox.ReadOnly = true;
            this.PreviewBox.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.PreviewBox.Size = new System.Drawing.Size(570, 148);
            this.PreviewBox.TabIndex = 9;
            this.PreviewBox.TabStop = false;
            this.PreviewBox.WordWrap = false;
            // 
            // PreviewLabel
            // 
            this.PreviewLabel.AutoSize = true;
            this.PreviewLabel.Location = new System.Drawing.Point(15, 190);
            this.PreviewLabel.Name = "PreviewLabel";
            this.PreviewLabel.Size = new System.Drawing.Size(48, 13);
            this.PreviewLabel.TabIndex = 8;
            this.PreviewLabel.Text = "Preview:";
            // 
            // MethodParametersGrid
            // 
            this.MethodParametersGrid.AllowUserToAddRows = false;
            this.MethodParametersGrid.AllowUserToDeleteRows = false;
            this.MethodParametersGrid.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right)));
            this.MethodParametersGrid.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.MethodParametersGrid.Location = new System.Drawing.Point(18, 82);
            this.MethodParametersGrid.Margin = new System.Windows.Forms.Padding(8, 3, 8, 3);
            this.MethodParametersGrid.Name = "MethodParametersGrid";
            this.MethodParametersGrid.Size = new System.Drawing.Size(570, 97);
            this.MethodParametersGrid.TabIndex = 2;
            // 
            // ParametersLabel
            // 
            this.ParametersLabel.AutoSize = true;
            this.ParametersLabel.Location = new System.Drawing.Point(15, 66);
            this.ParametersLabel.Name = "ParametersLabel";
            this.ParametersLabel.Size = new System.Drawing.Size(63, 13);
            this.ParametersLabel.TabIndex = 6;
            this.ParametersLabel.Text = "Parameters:";
            // 
            // MethodAccessibilityCombo
            // 
            this.MethodAccessibilityCombo.FormattingEnabled = true;
            this.MethodAccessibilityCombo.Location = new System.Drawing.Point(430, 34);
            this.MethodAccessibilityCombo.Name = "MethodAccessibilityCombo";
            this.MethodAccessibilityCombo.Size = new System.Drawing.Size(155, 21);
            this.MethodAccessibilityCombo.TabIndex = 1;
            // 
            // AccessibilityLabel
            // 
            this.AccessibilityLabel.AutoSize = true;
            this.AccessibilityLabel.Location = new System.Drawing.Point(356, 37);
            this.AccessibilityLabel.Name = "AccessibilityLabel";
            this.AccessibilityLabel.Size = new System.Drawing.Size(67, 13);
            this.AccessibilityLabel.TabIndex = 4;
            this.AccessibilityLabel.Text = "Accessibility:";
            // 
            // MethodNameBox
            // 
            this.MethodNameBox.Location = new System.Drawing.Point(83, 7);
            this.MethodNameBox.Name = "MethodNameBox";
            this.MethodNameBox.Size = new System.Drawing.Size(501, 20);
            this.MethodNameBox.TabIndex = 0;
            // 
            // NameLabel
            // 
            this.NameLabel.AutoSize = true;
            this.NameLabel.Location = new System.Drawing.Point(15, 10);
            this.NameLabel.Name = "NameLabel";
            this.NameLabel.Size = new System.Drawing.Size(38, 13);
            this.NameLabel.TabIndex = 0;
            this.NameLabel.Text = "Name:";
            // 
            // ExtractMethodDialog
            // 
            this.AcceptButton = this.OkButton;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.CancelButton = this.CancelDialogButton;
            this.ClientSize = new System.Drawing.Size(600, 481);
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.panel2);
            this.Controls.Add(this.flowLayoutPanel2);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "ExtractMethodDialog";
            this.ShowInTaskbar = false;
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
        private PictureBox InvalidNameValidationIcon;
        private TextBox PreviewBox;
        private Label PreviewLabel;
        private DataGridView MethodParametersGrid;
        private Label ParametersLabel;
        private ComboBox MethodAccessibilityCombo;
        private Label AccessibilityLabel;
        private TextBox MethodNameBox;
        private Label NameLabel;
    }
}
