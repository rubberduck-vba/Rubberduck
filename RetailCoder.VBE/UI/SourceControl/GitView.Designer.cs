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
            this.Fetch = new System.Windows.Forms.Button();
            this.NewBranch = new System.Windows.Forms.Button();
            this.Checkout = new System.Windows.Forms.Button();
            this.UserName = new System.Windows.Forms.TextBox();
            this.Password = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
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
            // Fetch
            // 
            this.Fetch.Location = new System.Drawing.Point(37, 124);
            this.Fetch.Name = "Fetch";
            this.Fetch.Size = new System.Drawing.Size(75, 23);
            this.Fetch.TabIndex = 3;
            this.Fetch.Text = "Fetch";
            this.Fetch.UseVisualStyleBackColor = true;
            this.Fetch.Click += new System.EventHandler(this.Fetch_Click);
            // 
            // NewBranch
            // 
            this.NewBranch.Location = new System.Drawing.Point(37, 154);
            this.NewBranch.Name = "NewBranch";
            this.NewBranch.Size = new System.Drawing.Size(75, 23);
            this.NewBranch.TabIndex = 4;
            this.NewBranch.Text = "New Branch";
            this.NewBranch.UseVisualStyleBackColor = true;
            this.NewBranch.Click += new System.EventHandler(this.NewBranch_Click);
            // 
            // Checkout
            // 
            this.Checkout.Location = new System.Drawing.Point(37, 184);
            this.Checkout.Name = "Checkout";
            this.Checkout.Size = new System.Drawing.Size(75, 23);
            this.Checkout.TabIndex = 5;
            this.Checkout.Text = "Checkout";
            this.Checkout.UseVisualStyleBackColor = true;
            this.Checkout.Click += new System.EventHandler(this.Checkout_Click);
            // 
            // UserName
            // 
            this.UserName.Location = new System.Drawing.Point(241, 21);
            this.UserName.Name = "UserName";
            this.UserName.Size = new System.Drawing.Size(80, 20);
            this.UserName.TabIndex = 6;
            // 
            // Password
            // 
            this.Password.Location = new System.Drawing.Point(241, 48);
            this.Password.Name = "Password";
            this.Password.PasswordChar = '*';
            this.Password.Size = new System.Drawing.Size(80, 20);
            this.Password.TabIndex = 7;
            this.Password.UseSystemPasswordChar = true;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(161, 27);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(60, 13);
            this.label1.TabIndex = 8;
            this.label1.Text = "User Name";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(164, 48);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(53, 13);
            this.label2.TabIndex = 9;
            this.label2.Text = "Password";
            // 
            // GitView
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(341, 455);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.Password);
            this.Controls.Add(this.UserName);
            this.Controls.Add(this.Checkout);
            this.Controls.Add(this.NewBranch);
            this.Controls.Add(this.Fetch);
            this.Controls.Add(this.Pull);
            this.Controls.Add(this.Push);
            this.Controls.Add(this.Commit);
            this.Name = "GitView";
            this.Text = "GitView";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button Commit;
        private System.Windows.Forms.Button Push;
        private System.Windows.Forms.Button Pull;
        private System.Windows.Forms.Button Fetch;
        private System.Windows.Forms.Button NewBranch;
        private System.Windows.Forms.Button Checkout;
        private System.Windows.Forms.TextBox UserName;
        private System.Windows.Forms.TextBox Password;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
    }
}