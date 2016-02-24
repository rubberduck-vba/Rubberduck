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
            this.DefaultRepositoryLocationTextBox = new System.Windows.Forms.TextBox();
            this.DefaultRepositoryLocationLabel = new System.Windows.Forms.Label();
            this.EmailAddressTextBox = new System.Windows.Forms.TextBox();
            this.EmailAddressLabel = new System.Windows.Forms.Label();
            this.UserNameTextBox = new System.Windows.Forms.TextBox();
            this.UserNameLabel = new System.Windows.Forms.Label();
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
            this.SettingsPanel.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.SettingsPanel.Name = "SettingsPanel";
            this.SettingsPanel.Padding = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.SettingsPanel.Size = new System.Drawing.Size(317, 609);
            this.SettingsPanel.TabIndex = 1;
            // 
            // RepositorySettingsBox
            // 
            this.RepositorySettingsBox.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.RepositorySettingsBox.Controls.Add(this.EditAttributeFileButton);
            this.RepositorySettingsBox.Controls.Add(this.EditIgnoreFileButton);
            this.RepositorySettingsBox.Location = new System.Drawing.Point(8, 261);
            this.RepositorySettingsBox.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.RepositorySettingsBox.Name = "RepositorySettingsBox";
            this.RepositorySettingsBox.Padding = new System.Windows.Forms.Padding(8, 7, 8, 7);
            this.RepositorySettingsBox.Size = new System.Drawing.Size(300, 87);
            this.RepositorySettingsBox.TabIndex = 3;
            this.RepositorySettingsBox.TabStop = false;
            this.RepositorySettingsBox.Text = "Repository Settings";
            // 
            // EditAttributeFileButton
            // 
            this.EditAttributeFileButton.Location = new System.Drawing.Point(164, 38);
            this.EditAttributeFileButton.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.EditAttributeFileButton.Name = "EditAttributeFileButton";
            this.EditAttributeFileButton.Size = new System.Drawing.Size(145, 28);
            this.EditAttributeFileButton.TabIndex = 1;
            this.EditAttributeFileButton.Text = "Attributes File";
            this.EditAttributeFileButton.UseVisualStyleBackColor = true;
            this.EditAttributeFileButton.Click += new System.EventHandler(this.EditAttributeButton_Click);
            // 
            // EditIgnoreFileButton
            // 
            this.EditIgnoreFileButton.Location = new System.Drawing.Point(9, 38);
            this.EditIgnoreFileButton.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.EditIgnoreFileButton.Name = "EditIgnoreFileButton";
            this.EditIgnoreFileButton.Size = new System.Drawing.Size(145, 28);
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
            this.GlobalSettingsBox.Controls.Add(this.DefaultRepositoryLocationTextBox);
            this.GlobalSettingsBox.Controls.Add(this.DefaultRepositoryLocationLabel);
            this.GlobalSettingsBox.Controls.Add(this.EmailAddressTextBox);
            this.GlobalSettingsBox.Controls.Add(this.EmailAddressLabel);
            this.GlobalSettingsBox.Controls.Add(this.UserNameTextBox);
            this.GlobalSettingsBox.Controls.Add(this.UserNameLabel);
            this.GlobalSettingsBox.Location = new System.Drawing.Point(8, 7);
            this.GlobalSettingsBox.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.GlobalSettingsBox.Name = "GlobalSettingsBox";
            this.GlobalSettingsBox.Padding = new System.Windows.Forms.Padding(8, 7, 8, 7);
            this.GlobalSettingsBox.Size = new System.Drawing.Size(300, 245);
            this.GlobalSettingsBox.TabIndex = 2;
            this.GlobalSettingsBox.TabStop = false;
            this.GlobalSettingsBox.Text = "Global Settings";
            // 
            // CancelGlobalSettingsButton
            // 
            this.CancelGlobalSettingsButton.Location = new System.Drawing.Point(164, 207);
            this.CancelGlobalSettingsButton.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.CancelGlobalSettingsButton.Name = "CancelGlobalSettingsButton";
            this.CancelGlobalSettingsButton.Size = new System.Drawing.Size(145, 28);
            this.CancelGlobalSettingsButton.TabIndex = 8;
            this.CancelGlobalSettingsButton.Text = "Cancel";
            this.CancelGlobalSettingsButton.UseVisualStyleBackColor = true;
            this.CancelGlobalSettingsButton.Click += new System.EventHandler(this.CancelGlobalSettingsButton_Click);
            // 
            // UpdateGlobalSettingsButton
            // 
            this.UpdateGlobalSettingsButton.Location = new System.Drawing.Point(9, 207);
            this.UpdateGlobalSettingsButton.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.UpdateGlobalSettingsButton.Name = "UpdateGlobalSettingsButton";
            this.UpdateGlobalSettingsButton.Size = new System.Drawing.Size(145, 28);
            this.UpdateGlobalSettingsButton.TabIndex = 7;
            this.UpdateGlobalSettingsButton.Text = "Update";
            this.UpdateGlobalSettingsButton.UseVisualStyleBackColor = true;
            this.UpdateGlobalSettingsButton.Click += new System.EventHandler(this.UpdateGlobalSettingsButton_Click);
            // 
            // BrowseDefaultRepositoryLocationButton
            // 
            this.BrowseDefaultRepositoryLocationButton.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.BrowseDefaultRepositoryLocationButton.Location = new System.Drawing.Point(241, 148);
            this.BrowseDefaultRepositoryLocationButton.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.BrowseDefaultRepositoryLocationButton.Name = "BrowseDefaultRepositoryLocationButton";
            this.BrowseDefaultRepositoryLocationButton.Size = new System.Drawing.Size(44, 25);
            this.BrowseDefaultRepositoryLocationButton.TabIndex = 6;
            this.BrowseDefaultRepositoryLocationButton.Text = "...";
            this.BrowseDefaultRepositoryLocationButton.UseVisualStyleBackColor = true;
            this.BrowseDefaultRepositoryLocationButton.Click += new System.EventHandler(this.BrowseDefaultRepositoryLocationButton_Click);
            // 
            // DefaultRepositoryLocationTextBox
            // 
            this.DefaultRepositoryLocationTextBox.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.DefaultRepositoryLocationTextBox.Location = new System.Drawing.Point(16, 148);
            this.DefaultRepositoryLocationTextBox.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.DefaultRepositoryLocationTextBox.Name = "DefaultRepositoryLocationTextBox";
            this.DefaultRepositoryLocationTextBox.Size = new System.Drawing.Size(225, 22);
            this.DefaultRepositoryLocationTextBox.TabIndex = 5;
            // 
            // DefaultRepositoryLocationLabel
            // 
            this.DefaultRepositoryLocationLabel.AutoSize = true;
            this.DefaultRepositoryLocationLabel.Location = new System.Drawing.Point(12, 127);
            this.DefaultRepositoryLocationLabel.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.DefaultRepositoryLocationLabel.Name = "DefaultRepositoryLocationLabel";
            this.DefaultRepositoryLocationLabel.Size = new System.Drawing.Size(183, 17);
            this.DefaultRepositoryLocationLabel.TabIndex = 4;
            this.DefaultRepositoryLocationLabel.Text = "Default Repository Location";
            // 
            // EmailAddressTextBox
            // 
            this.EmailAddressTextBox.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.EmailAddressTextBox.Location = new System.Drawing.Point(16, 98);
            this.EmailAddressTextBox.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.EmailAddressTextBox.Name = "EmailAddressTextBox";
            this.EmailAddressTextBox.Size = new System.Drawing.Size(269, 22);
            this.EmailAddressTextBox.TabIndex = 3;
            // 
            // EmailAddressLabel
            // 
            this.EmailAddressLabel.AutoSize = true;
            this.EmailAddressLabel.Location = new System.Drawing.Point(12, 78);
            this.EmailAddressLabel.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.EmailAddressLabel.Name = "EmailAddressLabel";
            this.EmailAddressLabel.Size = new System.Drawing.Size(98, 17);
            this.EmailAddressLabel.TabIndex = 2;
            this.EmailAddressLabel.Text = "Email Address";
            // 
            // UserNameTextBox
            // 
            this.UserNameTextBox.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.UserNameTextBox.Location = new System.Drawing.Point(17, 49);
            this.UserNameTextBox.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.UserNameTextBox.Name = "UserNameTextBox";
            this.UserNameTextBox.Size = new System.Drawing.Size(269, 22);
            this.UserNameTextBox.TabIndex = 1;
            // 
            // UserNameLabel
            // 
            this.UserNameLabel.AutoSize = true;
            this.UserNameLabel.Location = new System.Drawing.Point(13, 28);
            this.UserNameLabel.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.UserNameLabel.Name = "UserNameLabel";
            this.UserNameLabel.Size = new System.Drawing.Size(79, 17);
            this.UserNameLabel.TabIndex = 0;
            this.UserNameLabel.Text = "User Name";
            // 
            // SettingsControl
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.AutoSize = true;
            this.Controls.Add(this.SettingsPanel);
            this.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.Name = "SettingsControl";
            this.Size = new System.Drawing.Size(317, 609);
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
        private TextBox DefaultRepositoryLocationTextBox;
        private Label DefaultRepositoryLocationLabel;
        private TextBox EmailAddressTextBox;
        private Label EmailAddressLabel;
        private TextBox UserNameTextBox;
        private Label UserNameLabel;
    }
}
