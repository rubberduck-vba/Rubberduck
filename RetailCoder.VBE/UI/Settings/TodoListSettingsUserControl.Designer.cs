using System.ComponentModel;
using System.Windows.Forms;

namespace Rubberduck.UI.Settings
{
    partial class TodoListSettingsUserControl
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
            this.tokenListBox = new System.Windows.Forms.ListBox();
            this.tokenTextBox = new System.Windows.Forms.TextBox();
            this.priorityComboBox = new System.Windows.Forms.ComboBox();
            this.priorityLabel = new System.Windows.Forms.Label();
            this.tokenLabel = new System.Windows.Forms.Label();
            this.addButton = new System.Windows.Forms.Button();
            this.saveChangesButton = new System.Windows.Forms.Button();
            this.removeButton = new System.Windows.Forms.Button();
            this.tokenListLabel = new System.Windows.Forms.Label();
            this.InvalidNameValidationIcon = new System.Windows.Forms.PictureBox();
            ((System.ComponentModel.ISupportInitialize)(this.InvalidNameValidationIcon)).BeginInit();
            this.SuspendLayout();
            // 
            // tokenListBox
            // 
            this.tokenListBox.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left)));
            this.tokenListBox.FormattingEnabled = true;
            this.tokenListBox.ItemHeight = 16;
            this.tokenListBox.Location = new System.Drawing.Point(16, 32);
            this.tokenListBox.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.tokenListBox.Name = "tokenListBox";
            this.tokenListBox.Size = new System.Drawing.Size(177, 292);
            this.tokenListBox.TabIndex = 0;
            this.tokenListBox.SelectedIndexChanged += new System.EventHandler(this.tokenListBox_SelectedIndexChanged);
            // 
            // tokenTextBox
            // 
            this.tokenTextBox.Location = new System.Drawing.Point(203, 123);
            this.tokenTextBox.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.tokenTextBox.Name = "tokenTextBox";
            this.tokenTextBox.Size = new System.Drawing.Size(201, 22);
            this.tokenTextBox.TabIndex = 1;
            this.tokenTextBox.TextChanged += new System.EventHandler(this.tokenTextBox_TextChanged);
            // 
            // priorityComboBox
            // 
            this.priorityComboBox.FormattingEnabled = true;
            this.priorityComboBox.Location = new System.Drawing.Point(203, 52);
            this.priorityComboBox.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.priorityComboBox.Name = "priorityComboBox";
            this.priorityComboBox.Size = new System.Drawing.Size(201, 24);
            this.priorityComboBox.TabIndex = 2;
            this.priorityComboBox.SelectedIndexChanged += new System.EventHandler(this.priorityComboBox_SelectedIndexChanged);
            // 
            // priorityLabel
            // 
            this.priorityLabel.AutoSize = true;
            this.priorityLabel.Location = new System.Drawing.Point(199, 31);
            this.priorityLabel.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.priorityLabel.Name = "priorityLabel";
            this.priorityLabel.Size = new System.Drawing.Size(56, 17);
            this.priorityLabel.TabIndex = 3;
            this.priorityLabel.Text = "Priority:";
            // 
            // tokenLabel
            // 
            this.tokenLabel.AutoSize = true;
            this.tokenLabel.Location = new System.Drawing.Point(199, 102);
            this.tokenLabel.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.tokenLabel.Name = "tokenLabel";
            this.tokenLabel.Size = new System.Drawing.Size(87, 17);
            this.tokenLabel.TabIndex = 4;
            this.tokenLabel.Text = "Token Text: ";
            // 
            // addButton
            // 
            this.addButton.Location = new System.Drawing.Point(203, 178);
            this.addButton.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.addButton.Name = "addButton";
            this.addButton.Size = new System.Drawing.Size(203, 38);
            this.addButton.TabIndex = 5;
            this.addButton.Text = "Add";
            this.addButton.UseVisualStyleBackColor = true;
            this.addButton.Click += new System.EventHandler(this.addButton_Click);
            // 
            // saveChangesButton
            // 
            this.saveChangesButton.Location = new System.Drawing.Point(203, 233);
            this.saveChangesButton.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.saveChangesButton.Name = "saveChangesButton";
            this.saveChangesButton.Size = new System.Drawing.Size(203, 38);
            this.saveChangesButton.TabIndex = 6;
            this.saveChangesButton.Text = "Change";
            this.saveChangesButton.UseVisualStyleBackColor = true;
            this.saveChangesButton.Click += new System.EventHandler(this.saveChangesButton_Click);
            // 
            // removeButton
            // 
            this.removeButton.Location = new System.Drawing.Point(203, 289);
            this.removeButton.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.removeButton.Name = "removeButton";
            this.removeButton.Size = new System.Drawing.Size(203, 38);
            this.removeButton.TabIndex = 7;
            this.removeButton.Text = "Remove";
            this.removeButton.UseVisualStyleBackColor = true;
            this.removeButton.Click += new System.EventHandler(this.removeButton_Click);
            // 
            // tokenListLabel
            // 
            this.tokenListLabel.AutoSize = true;
            this.tokenListLabel.Location = new System.Drawing.Point(16, 5);
            this.tokenListLabel.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.tokenListLabel.Name = "tokenListLabel";
            this.tokenListLabel.Size = new System.Drawing.Size(78, 17);
            this.tokenListLabel.TabIndex = 8;
            this.tokenListLabel.Text = "Token List:";
            // 
            // InvalidNameValidationIcon
            // 
            this.InvalidNameValidationIcon.Image = global::Rubberduck.Properties.Resources.cross_circle;
            this.InvalidNameValidationIcon.Location = new System.Drawing.Point(395, 115);
            this.InvalidNameValidationIcon.Margin = new System.Windows.Forms.Padding(4);
            this.InvalidNameValidationIcon.Name = "InvalidNameValidationIcon";
            this.InvalidNameValidationIcon.Size = new System.Drawing.Size(16, 16);
            this.InvalidNameValidationIcon.SizeMode = System.Windows.Forms.PictureBoxSizeMode.AutoSize;
            this.InvalidNameValidationIcon.TabIndex = 14;
            this.InvalidNameValidationIcon.TabStop = false;
            // 
            // TodoListSettingsUserControl
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.InvalidNameValidationIcon);
            this.Controls.Add(this.tokenListLabel);
            this.Controls.Add(this.removeButton);
            this.Controls.Add(this.saveChangesButton);
            this.Controls.Add(this.addButton);
            this.Controls.Add(this.tokenLabel);
            this.Controls.Add(this.priorityLabel);
            this.Controls.Add(this.priorityComboBox);
            this.Controls.Add(this.tokenTextBox);
            this.Controls.Add(this.tokenListBox);
            this.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.MinimumSize = new System.Drawing.Size(419, 362);
            this.Name = "TodoListSettingsUserControl";
            this.Size = new System.Drawing.Size(419, 362);
            ((System.ComponentModel.ISupportInitialize)(this.InvalidNameValidationIcon)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private ListBox tokenListBox;
        private TextBox tokenTextBox;
        private ComboBox priorityComboBox;
        private Label priorityLabel;
        private Label tokenLabel;
        private Button addButton;
        private Button saveChangesButton;
        private Button removeButton;
        private Label tokenListLabel;
        private PictureBox InvalidNameValidationIcon;
    }
}
