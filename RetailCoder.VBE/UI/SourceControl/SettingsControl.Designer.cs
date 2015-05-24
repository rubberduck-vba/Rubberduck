using System.ComponentModel;
using System.Windows.Forms;

namespace Rubberduck.UI.SourceControl
{
    partial class SettingsControl
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
            this.SettingsPanel = new System.Windows.Forms.Panel();
            this.RepositorySettingsBox = new System.Windows.Forms.GroupBox();
            this.EditAttributeFileButton = new System.Windows.Forms.Button();
            this.EditIgnoreFileButton = new System.Windows.Forms.Button();
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
            this.RepositorySettingsBox.Controls.Add(this.EditAttributeFileButton);
            this.RepositorySettingsBox.Controls.Add(this.EditIgnoreFileButton);
            this.RepositorySettingsBox.Location = new System.Drawing.Point(6, 212);
            this.RepositorySettingsBox.Name = "RepositorySettingsBox";
            this.RepositorySettingsBox.Padding = new System.Windows.Forms.Padding(6);
            this.RepositorySettingsBox.Size = new System.Drawing.Size(225, 71);
            this.RepositorySettingsBox.TabIndex = 3;
            this.RepositorySettingsBox.TabStop = false;
            this.RepositorySettingsBox.Text = "Repository Settings";
            // 
            // EditAttributeFileButton
            // 
            this.EditAttributeFileButton.Location = new System.Drawing.Point(105, 31);
            this.EditAttributeFileButton.Name = "EditAttributeFileButton";
            this.EditAttributeFileButton.Size = new System.Drawing.Size(92, 23);
            this.EditAttributeFileButton.TabIndex = 1;
            this.EditAttributeFileButton.Text = "Attributes File";
            this.EditAttributeFileButton.UseVisualStyleBackColor = true;
            this.EditAttributeFileButton.Click += new System.EventHandler(this.EditAttributeButton_Click);
            // 
            // EditIgnoreFileButton
            // 
            this.EditIgnoreFileButton.Location = new System.Drawing.Point(7, 31);
            this.EditIgnoreFileButton.Name = "EditIgnoreFileButton";
            this.EditIgnoreFileButton.Size = new System.Drawing.Size(92, 23);
            this.EditIgnoreFileButton.TabIndex = 0;
            this.EditIgnoreFileButton.Text = "Ignore File";
            this.EditIgnoreFileButton.UseVisualStyleBackColor = true;
            this.EditIgnoreFileButton.Click += new System.EventHandler(this.EditIgnoreFileButton_Click);
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
            this.CancelGlobalSettingsButton.Click += new System.EventHandler(this.CancelGlobalSettingsButton_Click);
            // 
            // UpdateGlobalSettingsButton
            // 
            this.UpdateGlobalSettingsButton.Location = new System.Drawing.Point(7, 168);
            this.UpdateGlobalSettingsButton.Name = "UpdateGlobalSettingsButton";
            this.UpdateGlobalSettingsButton.Size = new System.Drawing.Size(92, 23);
            this.UpdateGlobalSettingsButton.TabIndex = 7;
            this.UpdateGlobalSettingsButton.Text = "Update";
            this.UpdateGlobalSettingsButton.UseVisualStyleBackColor = true;
            this.UpdateGlobalSettingsButton.Click += new System.EventHandler(this.UpdateGlobalSettingsButton_Click);
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
            this.BrowseDefaultRepositoryLocationButton.Click += new System.EventHandler(this.BrowseDefaultRepositoryLocationButton_Click);
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

        private Panel SettingsPanel;
        private GroupBox RepositorySettingsBox;
        private Button EditAttributeFileButton;
        private Button EditIgnoreFileButton;
        private GroupBox GlobalSettingsBox;
        private Button CancelGlobalSettingsButton;
        private Button UpdateGlobalSettingsButton;
        private Button BrowseDefaultRepositoryLocationButton;
        private TextBox DefaultRepositoryLocation;
        private Label label7;
        private TextBox EmailAddress;
        private Label label6;
        private TextBox UserName;
        private Label label5;
    }
}
