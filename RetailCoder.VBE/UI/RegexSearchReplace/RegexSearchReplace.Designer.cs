namespace Rubberduck.UI.RegexSearchReplace
{
    partial class RegexSearchReplace
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(RegexSearchReplace));
            this.flowLayoutPanel2 = new System.Windows.Forms.FlowLayoutPanel();
            this.CancelDialogButton = new System.Windows.Forms.Button();
            this.ReplaceAllButton = new System.Windows.Forms.Button();
            this.ReplaceButton = new System.Windows.Forms.Button();
            this.FindButton = new System.Windows.Forms.Button();
            this.FindBox = new System.Windows.Forms.TextBox();
            this.ReplaceBox = new System.Windows.Forms.TextBox();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.flowLayoutPanel2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
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
            this.flowLayoutPanel2.Location = new System.Drawing.Point(0, 47);
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
            // 
            // ReplaceAllButton
            // 
            this.ReplaceAllButton.BackColor = System.Drawing.SystemColors.ButtonFace;
            this.ReplaceAllButton.DialogResult = System.Windows.Forms.DialogResult.OK;
            this.ReplaceAllButton.Location = new System.Drawing.Point(229, 14);
            this.ReplaceAllButton.Margin = new System.Windows.Forms.Padding(4);
            this.ReplaceAllButton.Name = "ReplaceAllButton";
            this.ReplaceAllButton.Size = new System.Drawing.Size(100, 28);
            this.ReplaceAllButton.TabIndex = 1;
            this.ReplaceAllButton.Text = "Replace All";
            this.ReplaceAllButton.UseVisualStyleBackColor = false;
            // 
            // ReplaceButton
            // 
            this.ReplaceButton.BackColor = System.Drawing.SystemColors.ButtonFace;
            this.ReplaceButton.DialogResult = System.Windows.Forms.DialogResult.OK;
            this.ReplaceButton.Location = new System.Drawing.Point(121, 14);
            this.ReplaceButton.Margin = new System.Windows.Forms.Padding(4);
            this.ReplaceButton.Name = "ReplaceButton";
            this.ReplaceButton.Size = new System.Drawing.Size(100, 28);
            this.ReplaceButton.TabIndex = 2;
            this.ReplaceButton.Text = "Replace";
            this.ReplaceButton.UseVisualStyleBackColor = false;
            // 
            // FindButton
            // 
            this.FindButton.BackColor = System.Drawing.SystemColors.ButtonFace;
            this.FindButton.DialogResult = System.Windows.Forms.DialogResult.OK;
            this.FindButton.Location = new System.Drawing.Point(13, 14);
            this.FindButton.Margin = new System.Windows.Forms.Padding(4);
            this.FindButton.Name = "FindButton";
            this.FindButton.Size = new System.Drawing.Size(100, 28);
            this.FindButton.TabIndex = 3;
            this.FindButton.Text = "Find";
            this.FindButton.UseVisualStyleBackColor = false;
            // 
            // FindBox
            // 
            this.FindBox.Location = new System.Drawing.Point(13, 12);
            this.FindBox.Name = "FindBox";
            this.FindBox.Size = new System.Drawing.Size(194, 22);
            this.FindBox.TabIndex = 5;
            // 
            // ReplaceBox
            // 
            this.ReplaceBox.Location = new System.Drawing.Point(243, 12);
            this.ReplaceBox.Name = "ReplaceBox";
            this.ReplaceBox.Size = new System.Drawing.Size(194, 22);
            this.ReplaceBox.TabIndex = 6;
            // 
            // pictureBox1
            // 
            this.pictureBox1.Image = global::Rubberduck.Properties.Resources.arrow;
            this.pictureBox1.Location = new System.Drawing.Point(214, 12);
            this.pictureBox1.Margin = new System.Windows.Forms.Padding(4);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(22, 22);
            this.pictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.pictureBox1.TabIndex = 7;
            this.pictureBox1.TabStop = false;
            // 
            // RegexSearchReplace
            // 
            this.AcceptButton = this.FindButton;
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.CancelButton = this.CancelDialogButton;
            this.ClientSize = new System.Drawing.Size(452, 100);
            this.Controls.Add(this.ReplaceBox);
            this.Controls.Add(this.FindBox);
            this.Controls.Add(this.flowLayoutPanel2);
            this.Controls.Add(this.pictureBox1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "RegexSearchReplace";
            this.Text = "Regex Search & Replace";
            this.flowLayoutPanel2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.FlowLayoutPanel flowLayoutPanel2;
        private System.Windows.Forms.Button CancelDialogButton;
        private System.Windows.Forms.Button ReplaceAllButton;
        private System.Windows.Forms.Button ReplaceButton;
        private System.Windows.Forms.Button FindButton;
        private System.Windows.Forms.TextBox FindBox;
        private System.Windows.Forms.TextBox ReplaceBox;
        private System.Windows.Forms.PictureBox pictureBox1;
    }
}