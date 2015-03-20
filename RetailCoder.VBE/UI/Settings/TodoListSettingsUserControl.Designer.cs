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
            this.SuspendLayout();
            // 
            // tokenListBox
            // 
            this.tokenListBox.FormattingEnabled = true;
            this.tokenListBox.Location = new System.Drawing.Point(12, 26);
            this.tokenListBox.Name = "tokenListBox";
            this.tokenListBox.Size = new System.Drawing.Size(331, 238);
            this.tokenListBox.TabIndex = 0;
            this.tokenListBox.SelectedIndexChanged += new System.EventHandler(this.tokenListBox_SelectedIndexChanged);
            // 
            // tokenTextBox
            // 
            this.tokenTextBox.Location = new System.Drawing.Point(355, 98);
            this.tokenTextBox.Name = "tokenTextBox";
            this.tokenTextBox.Size = new System.Drawing.Size(152, 20);
            this.tokenTextBox.TabIndex = 1;
            this.tokenTextBox.TextChanged += new System.EventHandler(this.tokenTextBox_TextChanged);
            // 
            // priorityComboBox
            // 
            this.priorityComboBox.FormattingEnabled = true;
            this.priorityComboBox.Location = new System.Drawing.Point(355, 40);
            this.priorityComboBox.Name = "priorityComboBox";
            this.priorityComboBox.Size = new System.Drawing.Size(152, 21);
            this.priorityComboBox.TabIndex = 2;
            this.priorityComboBox.SelectedIndexChanged += new System.EventHandler(this.priorityComboBox_SelectedIndexChanged);
            // 
            // priorityLabel
            // 
            this.priorityLabel.AutoSize = true;
            this.priorityLabel.Location = new System.Drawing.Point(352, 23);
            this.priorityLabel.Name = "priorityLabel";
            this.priorityLabel.Size = new System.Drawing.Size(41, 13);
            this.priorityLabel.TabIndex = 3;
            this.priorityLabel.Text = "Priority:";
            // 
            // tokenLabel
            // 
            this.tokenLabel.AutoSize = true;
            this.tokenLabel.Location = new System.Drawing.Point(352, 81);
            this.tokenLabel.Name = "tokenLabel";
            this.tokenLabel.Size = new System.Drawing.Size(68, 13);
            this.tokenLabel.TabIndex = 4;
            this.tokenLabel.Text = "Token Text: ";
            // 
            // addButton
            // 
            this.addButton.Location = new System.Drawing.Point(355, 143);
            this.addButton.Name = "addButton";
            this.addButton.Size = new System.Drawing.Size(152, 31);
            this.addButton.TabIndex = 5;
            this.addButton.Text = "Add";
            this.addButton.UseVisualStyleBackColor = true;
            this.addButton.Click += new System.EventHandler(this.addButton_Click);
            // 
            // saveChangesButton
            // 
            this.saveChangesButton.Location = new System.Drawing.Point(355, 187);
            this.saveChangesButton.Name = "saveChangesButton";
            this.saveChangesButton.Size = new System.Drawing.Size(152, 31);
            this.saveChangesButton.TabIndex = 6;
            this.saveChangesButton.Text = "Change";
            this.saveChangesButton.UseVisualStyleBackColor = true;
            this.saveChangesButton.Click += new System.EventHandler(this.saveChangesButton_Click);
            // 
            // removeButton
            // 
            this.removeButton.Location = new System.Drawing.Point(355, 233);
            this.removeButton.Name = "removeButton";
            this.removeButton.Size = new System.Drawing.Size(152, 31);
            this.removeButton.TabIndex = 7;
            this.removeButton.Text = "Remove";
            this.removeButton.UseVisualStyleBackColor = true;
            this.removeButton.Click += new System.EventHandler(this.removeButton_Click);
            // 
            // tokenListLabel
            // 
            this.tokenListLabel.AutoSize = true;
            this.tokenListLabel.Location = new System.Drawing.Point(12, 4);
            this.tokenListLabel.Name = "tokenListLabel";
            this.tokenListLabel.Size = new System.Drawing.Size(60, 13);
            this.tokenListLabel.TabIndex = 8;
            this.tokenListLabel.Text = "Token List:";
            // 
            // TodoListSettingsControl
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.tokenListLabel);
            this.Controls.Add(this.removeButton);
            this.Controls.Add(this.saveChangesButton);
            this.Controls.Add(this.addButton);
            this.Controls.Add(this.tokenLabel);
            this.Controls.Add(this.priorityLabel);
            this.Controls.Add(this.priorityComboBox);
            this.Controls.Add(this.tokenTextBox);
            this.Controls.Add(this.tokenListBox);
            this.Name = "TodoListSettingsControl";
            this.Size = new System.Drawing.Size(530, 294);
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
    }
}
