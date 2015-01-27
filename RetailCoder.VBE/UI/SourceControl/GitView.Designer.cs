namespace Rubberduck.UI.SourceControl
{
    partial class GitView
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
            this.Commit = new System.Windows.Forms.Button();
            this.Push = new System.Windows.Forms.Button();
            this.Pull = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // Commit
            // 
            this.Commit.Location = new System.Drawing.Point(37, 34);
            this.Commit.Name = "Commit";
            this.Commit.Size = new System.Drawing.Size(75, 23);
            this.Commit.TabIndex = 0;
            this.Commit.Text = "Commit";
            this.Commit.UseVisualStyleBackColor = true;
            this.Commit.Click += new System.EventHandler(this.Commit_Click);
            // 
            // Push
            // 
            this.Push.Location = new System.Drawing.Point(37, 64);
            this.Push.Name = "Push";
            this.Push.Size = new System.Drawing.Size(75, 23);
            this.Push.TabIndex = 1;
            this.Push.Text = "Push";
            this.Push.UseVisualStyleBackColor = true;
            this.Push.Click += new System.EventHandler(this.Push_Click);
            // 
            // Pull
            // 
            this.Pull.Location = new System.Drawing.Point(37, 94);
            this.Pull.Name = "Pull";
            this.Pull.Size = new System.Drawing.Size(75, 23);
            this.Pull.TabIndex = 2;
            this.Pull.Text = "Pull";
            this.Pull.UseVisualStyleBackColor = true;
            this.Pull.Click += new System.EventHandler(this.Pull_Click);
            // 
            // GitView
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(341, 455);
            this.Controls.Add(this.Pull);
            this.Controls.Add(this.Push);
            this.Controls.Add(this.Commit);
            this.Name = "GitView";
            this.Text = "GitView";
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button Commit;
        private System.Windows.Forms.Button Push;
        private System.Windows.Forms.Button Pull;
    }
}