using System.ComponentModel;
using System.Windows.Forms;

namespace Rubberduck.UI
{
    partial class AboutWindow
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
            this.AppVersionLabel = new System.Windows.Forms.Label();
            this.SpecialThanksList = new System.Windows.Forms.ListBox();
            this.OfficialWebsiteLinkLabel = new System.Windows.Forms.LinkLabel();
            this.TwitterIcon = new System.Windows.Forms.PictureBox();
            this.FacebookIcon = new System.Windows.Forms.PictureBox();
            this.GitHubIcon = new System.Windows.Forms.PictureBox();
            this.GooglePlusIcon = new System.Windows.Forms.PictureBox();
            this.CloseButton = new System.Windows.Forms.Button();
            this.CopyrightLabel = new System.Windows.Forms.Label();
            this.AttributionsLabel = new System.Windows.Forms.Label();
            this.SpecialThanksLabel = new System.Windows.Forms.Label();
            this.AttributionsList = new System.Windows.Forms.ListBox();
            ((System.ComponentModel.ISupportInitialize)(this.TwitterIcon)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.FacebookIcon)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.GitHubIcon)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.GooglePlusIcon)).BeginInit();
            this.SuspendLayout();
            // 
            // AppVersionLabel
            // 
            this.AppVersionLabel.AutoSize = true;
            this.AppVersionLabel.BackColor = System.Drawing.Color.Transparent;
            this.AppVersionLabel.Font = new System.Drawing.Font("Microsoft Sans Serif", 14F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.AppVersionLabel.Location = new System.Drawing.Point(180, 13);
            this.AppVersionLabel.Name = "AppVersionLabel";
            this.AppVersionLabel.Size = new System.Drawing.Size(90, 24);
            this.AppVersionLabel.TabIndex = 0;
            this.AppVersionLabel.Text = "[version]";
            // 
            // SpecialThanksList
            // 
            this.SpecialThanksList.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.SpecialThanksList.ForeColor = System.Drawing.Color.DimGray;
            this.SpecialThanksList.FormattingEnabled = true;
            this.SpecialThanksList.Items.AddRange(new object[] {
            global::Rubberduck.UI.RubberduckUI.About_Community,
            "Code Review Stack Exchange",
            "JetBrains ReSharper Community Team",
            "Stack Overflow",
            "",
            global::Rubberduck.UI.RubberduckUI.About_Blogs,
            "Michal Krzych (vba4all.com)",
            "Knjname developer blog (clockahead.blogspot.jp)",
            "",
            global::Rubberduck.UI.RubberduckUI.About_Contributors,
            "Abraham Hosch",
            "Carlos J. Quintero (MZ-Tools articles & help with VBE API)",
            "@daFreeMan",
            "@Duga SE chat bot",
            "Francis Veilleux-Gaboury",
            "Frank Van Heeswijk",
            "@Heslacher",
            "Jeroen Vannevel dos Sànchez di Castello du Aragon de Pompidou",
            "@mjolka",
            "Philip Wales",
            "Rob Bovey",
            "Ross McLean",
            "Ross Knudsen",
            "Simon Forsberg",
            "Stephen Bullen",
            "",
            global::Rubberduck.UI.RubberduckUI.About_AllContributors,
            global::Rubberduck.UI.RubberduckUI.About_Stargazers,
            global::Rubberduck.UI.RubberduckUI.About_Anyone});
            this.SpecialThanksList.Location = new System.Drawing.Point(187, 164);
            this.SpecialThanksList.Name = "SpecialThanksList";
            this.SpecialThanksList.SelectionMode = System.Windows.Forms.SelectionMode.None;
            this.SpecialThanksList.Size = new System.Drawing.Size(300, 104);
            this.SpecialThanksList.TabIndex = 1;
            // 
            // OfficialWebsiteLinkLabel
            // 
            this.OfficialWebsiteLinkLabel.ActiveLinkColor = System.Drawing.Color.DimGray;
            this.OfficialWebsiteLinkLabel.AutoSize = true;
            this.OfficialWebsiteLinkLabel.BackColor = System.Drawing.Color.Transparent;
            this.OfficialWebsiteLinkLabel.Cursor = System.Windows.Forms.Cursors.Hand;
            this.OfficialWebsiteLinkLabel.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.OfficialWebsiteLinkLabel.ForeColor = System.Drawing.Color.DimGray;
            this.OfficialWebsiteLinkLabel.LinkBehavior = System.Windows.Forms.LinkBehavior.HoverUnderline;
            this.OfficialWebsiteLinkLabel.LinkColor = System.Drawing.Color.DimGray;
            this.OfficialWebsiteLinkLabel.Location = new System.Drawing.Point(12, 126);
            this.OfficialWebsiteLinkLabel.Name = "OfficialWebsiteLinkLabel";
            this.OfficialWebsiteLinkLabel.Size = new System.Drawing.Size(156, 17);
            this.OfficialWebsiteLinkLabel.TabIndex = 2;
            this.OfficialWebsiteLinkLabel.TabStop = true;
            this.OfficialWebsiteLinkLabel.Text = "rubberduck-vba.com";
            this.OfficialWebsiteLinkLabel.TextAlign = System.Drawing.ContentAlignment.TopCenter;
            // 
            // TwitterIcon
            // 
            this.TwitterIcon.Cursor = System.Windows.Forms.Cursors.Hand;
            this.TwitterIcon.Image = global::Rubberduck.Properties.Resources.twitter_circle_black_512;
            this.TwitterIcon.Location = new System.Drawing.Point(53, 266);
            this.TwitterIcon.Name = "TwitterIcon";
            this.TwitterIcon.Size = new System.Drawing.Size(32, 32);
            this.TwitterIcon.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.TwitterIcon.TabIndex = 3;
            this.TwitterIcon.TabStop = false;
            // 
            // FacebookIcon
            // 
            this.FacebookIcon.Cursor = System.Windows.Forms.Cursors.Hand;
            this.FacebookIcon.Image = global::Rubberduck.Properties.Resources.facebook_circle_256;
            this.FacebookIcon.Location = new System.Drawing.Point(91, 267);
            this.FacebookIcon.Name = "FacebookIcon";
            this.FacebookIcon.Size = new System.Drawing.Size(30, 30);
            this.FacebookIcon.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.FacebookIcon.TabIndex = 3;
            this.FacebookIcon.TabStop = false;
            // 
            // GitHubIcon
            // 
            this.GitHubIcon.Cursor = System.Windows.Forms.Cursors.Hand;
            this.GitHubIcon.Image = global::Rubberduck.Properties.Resources.github_circle_black_128;
            this.GitHubIcon.Location = new System.Drawing.Point(15, 266);
            this.GitHubIcon.Name = "GitHubIcon";
            this.GitHubIcon.Size = new System.Drawing.Size(32, 32);
            this.GitHubIcon.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.GitHubIcon.TabIndex = 3;
            this.GitHubIcon.TabStop = false;
            // 
            // GooglePlusIcon
            // 
            this.GooglePlusIcon.Cursor = System.Windows.Forms.Cursors.Hand;
            this.GooglePlusIcon.Image = global::Rubberduck.Properties.Resources.google_circle_512;
            this.GooglePlusIcon.Location = new System.Drawing.Point(129, 267);
            this.GooglePlusIcon.Name = "GooglePlusIcon";
            this.GooglePlusIcon.Size = new System.Drawing.Size(30, 30);
            this.GooglePlusIcon.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.GooglePlusIcon.TabIndex = 3;
            this.GooglePlusIcon.TabStop = false;
            // 
            // CloseButton
            // 
            this.CloseButton.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.CloseButton.Location = new System.Drawing.Point(46, 193);
            this.CloseButton.Name = "CloseButton";
            this.CloseButton.Size = new System.Drawing.Size(0, 0);
            this.CloseButton.TabIndex = 4;
            this.CloseButton.Text = "Close";
            this.CloseButton.UseVisualStyleBackColor = true;
            // 
            // CopyrightLabel
            // 
            this.CopyrightLabel.AutoSize = true;
            this.CopyrightLabel.BackColor = System.Drawing.Color.Transparent;
            this.CopyrightLabel.Font = new System.Drawing.Font("Arial Narrow", 8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.CopyrightLabel.ForeColor = System.Drawing.Color.DimGray;
            this.CopyrightLabel.Location = new System.Drawing.Point(181, 282);
            this.CopyrightLabel.Name = "CopyrightLabel";
            this.CopyrightLabel.Size = new System.Drawing.Size(278, 15);
            this.CopyrightLabel.TabIndex = 5;
            this.CopyrightLabel.Text = "© Copyright 2014-2015 Mathieu Guindon & Christopher McClellan";
            this.CopyrightLabel.UseMnemonic = false;
            // 
            // AttributionsLabel
            // 
            this.AttributionsLabel.AutoSize = true;
            this.AttributionsLabel.BackColor = System.Drawing.Color.Transparent;
            this.AttributionsLabel.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.AttributionsLabel.Location = new System.Drawing.Point(184, 54);
            this.AttributionsLabel.Name = "AttributionsLabel";
            this.AttributionsLabel.Size = new System.Drawing.Size(91, 17);
            this.AttributionsLabel.TabIndex = 6;
            this.AttributionsLabel.Text = "Attributions";
            // 
            // SpecialThanksLabel
            // 
            this.SpecialThanksLabel.AutoSize = true;
            this.SpecialThanksLabel.BackColor = System.Drawing.Color.Transparent;
            this.SpecialThanksLabel.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.SpecialThanksLabel.Location = new System.Drawing.Point(184, 144);
            this.SpecialThanksLabel.Name = "SpecialThanksLabel";
            this.SpecialThanksLabel.Size = new System.Drawing.Size(119, 17);
            this.SpecialThanksLabel.TabIndex = 6;
            this.SpecialThanksLabel.Text = "Special Thanks";
            // 
            // AttributionsList
            // 
            this.AttributionsList.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.AttributionsList.ForeColor = System.Drawing.Color.DimGray;
            this.AttributionsList.FormattingEnabled = true;
            this.AttributionsList.Items.AddRange(new object[] {
            global::Rubberduck.UI.RubberduckUI.About_ParsingCredit,
            global::Rubberduck.UI.RubberduckUI.About_LibGit2SharpCredit,
            global::Rubberduck.UI.RubberduckUI.About_FugueIconCredit});
            this.AttributionsList.Location = new System.Drawing.Point(187, 74);
            this.AttributionsList.Name = "AttributionsList";
            this.AttributionsList.SelectionMode = System.Windows.Forms.SelectionMode.None;
            this.AttributionsList.Size = new System.Drawing.Size(288, 52);
            this.AttributionsList.TabIndex = 1;
            // 
            // AboutWindow
            // 
            this.AcceptButton = this.CloseButton;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackgroundImage = global::Rubberduck.Properties.Resources.RD_AboutWindow;
            this.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None;
            this.CancelButton = this.CloseButton;
            this.ClientSize = new System.Drawing.Size(499, 310);
            this.Controls.Add(this.SpecialThanksLabel);
            this.Controls.Add(this.AttributionsLabel);
            this.Controls.Add(this.CopyrightLabel);
            this.Controls.Add(this.CloseButton);
            this.Controls.Add(this.GitHubIcon);
            this.Controls.Add(this.GooglePlusIcon);
            this.Controls.Add(this.FacebookIcon);
            this.Controls.Add(this.TwitterIcon);
            this.Controls.Add(this.OfficialWebsiteLinkLabel);
            this.Controls.Add(this.AttributionsList);
            this.Controls.Add(this.SpecialThanksList);
            this.Controls.Add(this.AppVersionLabel);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "AboutWindow";
            this.ShowIcon = false;
            this.ShowInTaskbar = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "About Rubberduck";
            this.TopMost = true;
            ((System.ComponentModel.ISupportInitialize)(this.TwitterIcon)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.FacebookIcon)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.GitHubIcon)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.GooglePlusIcon)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private Label AppVersionLabel;
        private ListBox SpecialThanksList;
        private LinkLabel OfficialWebsiteLinkLabel;
        private PictureBox TwitterIcon;
        private PictureBox FacebookIcon;
        private PictureBox GitHubIcon;
        private PictureBox GooglePlusIcon;
        private Button CloseButton;
        private Label CopyrightLabel;
        private Label AttributionsLabel;
        private Label SpecialThanksLabel;
        private ListBox AttributionsList;

    }
}