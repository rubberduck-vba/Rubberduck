namespace Rubberduck.UI.SourceControl
{
    partial class SettingsControl
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

        #region Component Designer generated code

        /// <summary> 
        /// Required method for Designer support - do not modify 
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.SettingsPanel = new System.Windows.Forms.Panel();
            this.RepositorySettingsBox = new System.Windows.Forms.GroupBox();
            this.button2 = new System.Windows.Forms.Button();
            this.button1 = new System.Windows.Forms.Button();
            this.GlobalSettingsBox = new System.Windows.Forms.GroupBox();
            this.CancelGlobalSettingsButton = new System.Windows.Forms.Button();
            this.UpdateGlobalSettingsButton = new System.Windows.Forms.Button();
            this.BrowseDefaultRepositoryLocationButton = new System.Windows.Forms.Button();
            this.DefaultRepositoryLocation = new System.Windows.Forms.TextBox();
            this.label7 = new System.Windows.Forms.Label();
            this.EmailAddress = new System.Windows.Forms.TextBox();
            this.label6 = new System.Windows.Forms.Label();
            this.UserName = new System.Windows.Forms.TextBox();
            this.label5 = new System.Windows.Forms.Label();
            this.SettingsPanel.SuspendLayout();
            this.RepositorySettingsBox.SuspendLayout();
            this.GlobalSettingsBox.SuspendLayout();
            this.SuspendLayout();
            // 
            // SettingsPanel
            // 
            this.SettingsPanel.AutoScroll = true;
            this.SettingsPanel.Controls.Add(this.RepositorySettingsBox);
            this.SettingsPanel.Controls.Add(this.GlobalSettingsBox);
            this.SettingsPanel.Dock = System.Windows.Forms.DockStyle.Fill;
            this.SettingsPanel.Location = new System.Drawing.Point(0, 0);
            this.SettingsPanel.Name = "SettingsPanel";
            this.SettingsPanel.Padding = new System.Windows.Forms.Padding(3);
            this.SettingsPanel.Size = new System.Drawing.Size(238, 495);
            this.SettingsPanel.TabIndex = 1;
            // 
            // RepositorySettingsBox
            // 
            this.RepositorySettingsBox.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.RepositorySettingsBox.Controls.Add(this.button2);
            this.RepositorySettingsBox.Controls.Add(this.button1);
            this.RepositorySettingsBox.Location = new System.Drawing.Point(6, 212);
            this.RepositorySettingsBox.Name = "RepositorySettingsBox";
            this.RepositorySettingsBox.Padding = new System.Windows.Forms.Padding(6);
            this.RepositorySettingsBox.Size = new System.Drawing.Size(225, 71);
            this.RepositorySettingsBox.TabIndex = 3;
            this.RepositorySettingsBox.TabStop = false;
            this.RepositorySettingsBox.Text = "Repository Settings";
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(105, 31);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(92, 23);
            this.button2.TabIndex = 1;
            this.button2.Text = "Attributes File";
            this.button2.UseVisualStyleBackColor = true;
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(7, 31);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(92, 23);
            this.button1.TabIndex = 0;
            this.button1.Text = "Ignore File";
            this.button1.UseVisualStyleBackColor = true;
            // 
            // GlobalSettingsBox
            // 
            this.GlobalSettingsBox.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.GlobalSettingsBox.Controls.Add(this.CancelGlobalSettingsButton);
            this.GlobalSettingsBox.Controls.Add(this.UpdateGlobalSettingsButton);
            this.GlobalSettingsBox.Controls.Add(this.BrowseDefaultRepositoryLocationButton);
            this.GlobalSettingsBox.Controls.Add(this.DefaultRepositoryLocation);
            this.GlobalSettingsBox.Controls.Add(this.label7);
            this.GlobalSettingsBox.Controls.Add(this.EmailAddress);
            this.GlobalSettingsBox.Controls.Add(this.label6);
            this.GlobalSettingsBox.Controls.Add(this.UserName);
            this.GlobalSettingsBox.Controls.Add(this.label5);
            this.GlobalSettingsBox.Location = new System.Drawing.Point(6, 6);
            this.GlobalSettingsBox.Name = "GlobalSettingsBox";
            this.GlobalSettingsBox.Padding = new System.Windows.Forms.Padding(6);
            this.GlobalSettingsBox.Size = new System.Drawing.Size(225, 199);
            this.GlobalSettingsBox.TabIndex = 2;
            this.GlobalSettingsBox.TabStop = false;
            this.GlobalSettingsBox.Text = "Global Settings";
            // 
            // CancelGlobalSettingsButton
            // 
            this.CancelGlobalSettingsButton.Location = new System.Drawing.Point(105, 168);
            this.CancelGlobalSettingsButton.Name = "CancelGlobalSettingsButton";
            this.CancelGlobalSettingsButton.Size = new System.Drawing.Size(92, 23);
            this.CancelGlobalSettingsButton.TabIndex = 8;
            this.CancelGlobalSettingsButton.Text = "Cancel";
            this.CancelGlobalSettingsButton.UseVisualStyleBackColor = true;
            // 
            // UpdateGlobalSettingsButton
            // 
            this.UpdateGlobalSettingsButton.Location = new System.Drawing.Point(7, 168);
            this.UpdateGlobalSettingsButton.Name = "UpdateGlobalSettingsButton";
            this.UpdateGlobalSettingsButton.Size = new System.Drawing.Size(92, 23);
            this.UpdateGlobalSettingsButton.TabIndex = 7;
            this.UpdateGlobalSettingsButton.Text = "Update";
            this.UpdateGlobalSettingsButton.UseVisualStyleBackColor = true;
            // 
            // BrowseDefaultRepositoryLocationButton
            // 
            this.BrowseDefaultRepositoryLocationButton.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.BrowseDefaultRepositoryLocationButton.Location = new System.Drawing.Point(181, 120);
            this.BrowseDefaultRepositoryLocationButton.Name = "BrowseDefaultRepositoryLocationButton";
            this.BrowseDefaultRepositoryLocationButton.Size = new System.Drawing.Size(33, 20);
            this.BrowseDefaultRepositoryLocationButton.TabIndex = 6;
            this.BrowseDefaultRepositoryLocationButton.Text = "...";
            this.BrowseDefaultRepositoryLocationButton.UseVisualStyleBackColor = true;
            // 
            // DefaultRepositoryLocation
            // 
            this.DefaultRepositoryLocation.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.DefaultRepositoryLocation.Location = new System.Drawing.Point(12, 120);
            this.DefaultRepositoryLocation.Name = "DefaultRepositoryLocation";
            this.DefaultRepositoryLocation.Size = new System.Drawing.Size(170, 20);
            this.DefaultRepositoryLocation.TabIndex = 5;
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(9, 103);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(138, 13);
            this.label7.TabIndex = 4;
            this.label7.Text = "Default Repository Location";
            // 
            // EmailAddress
            // 
            this.EmailAddress.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.EmailAddress.Location = new System.Drawing.Point(12, 80);
            this.EmailAddress.Name = "EmailAddress";
            this.EmailAddress.Size = new System.Drawing.Size(203, 20);
            this.EmailAddress.TabIndex = 3;
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(9, 63);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(73, 13);
            this.label6.TabIndex = 2;
            this.label6.Text = "Email Address";
            // 
            // UserName
            // 
            this.UserName.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.UserName.Location = new System.Drawing.Point(13, 40);
            this.UserName.Name = "UserName";
            this.UserName.Size = new System.Drawing.Size(203, 20);
            this.UserName.TabIndex = 1;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(10, 23);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(60, 13);
            this.label5.TabIndex = 0;
            this.label5.Text = "User Name";
            // 
            // SettingsControl
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.AutoSize = true;
            this.Controls.Add(this.SettingsPanel);
            this.Name = "SettingsControl";
            this.Size = new System.Drawing.Size(238, 495);
            this.SettingsPanel.ResumeLayout(false);
            this.RepositorySettingsBox.ResumeLayout(false);
            this.GlobalSettingsBox.ResumeLayout(false);
            this.GlobalSettingsBox.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Panel SettingsPanel;
        private System.Windows.Forms.GroupBox RepositorySettingsBox;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.GroupBox GlobalSettingsBox;
        private System.Windows.Forms.Button CancelGlobalSettingsButton;
        private System.Windows.Forms.Button UpdateGlobalSettingsButton;
        private System.Windows.Forms.Button BrowseDefaultRepositoryLocationButton;
        private System.Windows.Forms.TextBox DefaultRepositoryLocation;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.TextBox EmailAddress;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.TextBox UserName;
        private System.Windows.Forms.Label label5;
    }
}
