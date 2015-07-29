namespace Rubberduck.UI.RegexSearchReplace
{
    partial class RegexSearchReplaceDialog
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(RegexSearchReplaceDialog));
            this.flowLayoutPanel2 = new System.Windows.Forms.FlowLayoutPanel();
            this.CancelDialogButton = new System.Windows.Forms.Button();
            this.ReplaceAllButton = new System.Windows.Forms.Button();
            this.ReplaceButton = new System.Windows.Forms.Button();
            this.FindButton = new System.Windows.Forms.Button();
            this.ReplaceBox = new System.Windows.Forms.TextBox();
            this.SearchBox = new System.Windows.Forms.TextBox();
            this.SearchLabel = new System.Windows.Forms.Label();
            this.ReplaceLabel = new System.Windows.Forms.Label();
            this.ScopeComboBox = new System.Windows.Forms.ComboBox();
            this.ScopeLabel = new System.Windows.Forms.Label();
            this.flowLayoutPanel2.SuspendLayout();
            this.SuspendLayout();
            // 
            // flowLayoutPanel2
            // 
            this.flowLayoutPanel2.BackColor = System.Drawing.SystemColors.ControlDark;
            this.flowLayoutPanel2.Controls.Add(this.CancelDialogButton);
            this.flowLayoutPanel2.Controls.Add(this.ReplaceAllButton);
            this.flowLayoutPanel2.Controls.Add(this.ReplaceButton);
            this.flowLayoutPanel2.Controls.Add(this.FindButton);
            this.flowLayoutPanel2.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.flowLayoutPanel2.FlowDirection = System.Windows.Forms.FlowDirection.RightToLeft;
            this.flowLayoutPanel2.Location = new System.Drawing.Point(0, 141);
            this.flowLayoutPanel2.Margin = new System.Windows.Forms.Padding(4);
            this.flowLayoutPanel2.Name = "flowLayoutPanel2";
            this.flowLayoutPanel2.Padding = new System.Windows.Forms.Padding(11, 10, 0, 10);
            this.flowLayoutPanel2.Size = new System.Drawing.Size(452, 53);
            this.flowLayoutPanel2.TabIndex = 4;
            // 
            // CancelDialogButton
            // 
            this.CancelDialogButton.BackColor = System.Drawing.SystemColors.ButtonFace;
            this.CancelDialogButton.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.CancelDialogButton.Location = new System.Drawing.Point(337, 14);
            this.CancelDialogButton.Margin = new System.Windows.Forms.Padding(4);
            this.CancelDialogButton.Name = "CancelDialogButton";
            this.CancelDialogButton.Size = new System.Drawing.Size(100, 28);
            this.CancelDialogButton.TabIndex = 0;
            this.CancelDialogButton.Text = "Cancel";
            this.CancelDialogButton.UseVisualStyleBackColor = false;
            this.CancelDialogButton.Click += new System.EventHandler(this.OnCancelButtonClicked);
            // 
            // ReplaceAllButton
            // 
            this.ReplaceAllButton.BackColor = System.Drawing.SystemColors.ButtonFace;
            this.ReplaceAllButton.Location = new System.Drawing.Point(229, 14);
            this.ReplaceAllButton.Margin = new System.Windows.Forms.Padding(4);
            this.ReplaceAllButton.Name = "ReplaceAllButton";
            this.ReplaceAllButton.Size = new System.Drawing.Size(100, 28);
            this.ReplaceAllButton.TabIndex = 1;
            this.ReplaceAllButton.Text = "Replace All";
            this.ReplaceAllButton.UseVisualStyleBackColor = false;
            this.ReplaceAllButton.Click += new System.EventHandler(this.OnReplaceAllButtonClicked);
            // 
            // ReplaceButton
            // 
            this.ReplaceButton.BackColor = System.Drawing.SystemColors.ButtonFace;
            this.ReplaceButton.Location = new System.Drawing.Point(121, 14);
            this.ReplaceButton.Margin = new System.Windows.Forms.Padding(4);
            this.ReplaceButton.Name = "ReplaceButton";
            this.ReplaceButton.Size = new System.Drawing.Size(100, 28);
            this.ReplaceButton.TabIndex = 2;
            this.ReplaceButton.Text = "Replace";
            this.ReplaceButton.UseVisualStyleBackColor = false;
            this.ReplaceButton.Click += new System.EventHandler(this.OnReplaceButtonClicked);
            // 
            // FindButton
            // 
            this.FindButton.BackColor = System.Drawing.SystemColors.ButtonFace;
            this.FindButton.Location = new System.Drawing.Point(13, 14);
            this.FindButton.Margin = new System.Windows.Forms.Padding(4);
            this.FindButton.Name = "FindButton";
            this.FindButton.Size = new System.Drawing.Size(100, 28);
            this.FindButton.TabIndex = 3;
            this.FindButton.Text = "Find";
            this.FindButton.UseVisualStyleBackColor = false;
            this.FindButton.Click += new System.EventHandler(this.OnFindButtonClicked);
            // 
            // ReplaceBox
            // 
            this.ReplaceBox.Location = new System.Drawing.Point(107, 52);
            this.ReplaceBox.Name = "ReplaceBox";
            this.ReplaceBox.Size = new System.Drawing.Size(330, 22);
            this.ReplaceBox.TabIndex = 5;
            // 
            // SearchBox
            // 
            this.SearchBox.Location = new System.Drawing.Point(107, 12);
            this.SearchBox.Name = "SearchBox";
            this.SearchBox.Size = new System.Drawing.Size(330, 22);
            this.SearchBox.TabIndex = 6;
            // 
            // SearchLabel
            // 
            this.SearchLabel.AutoSize = true;
            this.SearchLabel.Location = new System.Drawing.Point(10, 12);
            this.SearchLabel.Name = "SearchLabel";
            this.SearchLabel.Size = new System.Drawing.Size(57, 17);
            this.SearchLabel.TabIndex = 7;
            this.SearchLabel.Text = "Search:";
            // 
            // ReplaceLabel
            // 
            this.ReplaceLabel.AutoSize = true;
            this.ReplaceLabel.Location = new System.Drawing.Point(10, 52);
            this.ReplaceLabel.Name = "ReplaceLabel";
            this.ReplaceLabel.Size = new System.Drawing.Size(64, 17);
            this.ReplaceLabel.TabIndex = 8;
            this.ReplaceLabel.Text = "Replace:";
            // 
            // ScopeComboBox
            // 
            this.ScopeComboBox.FormattingEnabled = true;
            this.ScopeComboBox.Location = new System.Drawing.Point(107, 92);
            this.ScopeComboBox.Name = "ScopeComboBox";
            this.ScopeComboBox.Size = new System.Drawing.Size(330, 24);
            this.ScopeComboBox.TabIndex = 9;
            // 
            // ScopeLabel
            // 
            this.ScopeLabel.AutoSize = true;
            this.ScopeLabel.Location = new System.Drawing.Point(10, 92);
            this.ScopeLabel.Name = "ScopeLabel";
            this.ScopeLabel.Size = new System.Drawing.Size(52, 17);
            this.ScopeLabel.TabIndex = 10;
            this.ScopeLabel.Text = "Scope:";
            // 
            // RegexSearchReplaceDialog
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.AutoValidate = System.Windows.Forms.AutoValidate.Disable;
            this.CancelButton = this.CancelDialogButton;
            this.ClientSize = new System.Drawing.Size(452, 194);
            this.Controls.Add(this.ScopeLabel);
            this.Controls.Add(this.ScopeComboBox);
            this.Controls.Add(this.ReplaceLabel);
            this.Controls.Add(this.SearchLabel);
            this.Controls.Add(this.SearchBox);
            this.Controls.Add(this.ReplaceBox);
            this.Controls.Add(this.flowLayoutPanel2);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "RegexSearchReplaceDialog";
            this.Text = "Regex Search & Replace";
            this.flowLayoutPanel2.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.FlowLayoutPanel flowLayoutPanel2;
        private System.Windows.Forms.Button CancelDialogButton;
        private System.Windows.Forms.Button ReplaceAllButton;
        private System.Windows.Forms.Button ReplaceButton;
        private System.Windows.Forms.Button FindButton;
        private System.Windows.Forms.TextBox ReplaceBox;
        private System.Windows.Forms.TextBox SearchBox;
        private System.Windows.Forms.Label SearchLabel;
        private System.Windows.Forms.Label ReplaceLabel;
        private System.Windows.Forms.ComboBox ScopeComboBox;
        private System.Windows.Forms.Label ScopeLabel;
    }
}