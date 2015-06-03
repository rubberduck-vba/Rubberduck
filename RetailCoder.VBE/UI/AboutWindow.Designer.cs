using System.ComponentModel;
using System.Windows.Forms;

namespace Rubberduck.UI
{
    partial class _AboutWindow
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
            this.AppVersionLabel.Location = new System.Drawing.Point(240, 16);
            this.AppVersionLabel.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.AppVersionLabel.Name = "AppVersionLabel";
            this.AppVersionLabel.Size = new System.Drawing.Size(114, 29);
            this.AppVersionLabel.TabIndex = 0;
            this.AppVersionLabel.Text = "[version]";
            // 
            // SpecialThanksList
            // 
            this.SpecialThanksList.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.SpecialThanksList.ForeColor = System.Drawing.Color.DimGray;
            this.SpecialThanksList.FormattingEnabled = true;
            this.SpecialThanksList.ItemHeight = 16;
            this.SpecialThanksList.Items.AddRange(new object[] {
            "Community:",
            "Code Review Stack Exchange",
            "JetBrains ReSharper Community Team",
            "Stack Overflow",
            "",
            "Blogs:",
            "Michal Krzych (vba4all.com)",
            "Knjname developer blog (clockahead.blogspot.jp)",
            "",
            "Contributors & supporters:",
            "Abraham Hosch",
            "Carlos J. Quintero (MZ-Tools articles & help with VBE API)",
            "@daFreeMan",
            "@Duga SE chat bot",
            "Francis Veilleux-Gaboury",
            "Frank Van Heeswijk",
            "@mjolka",
            "Philip Wales",
            "Rob Bovey",
            "Ross McLean",
            "Ross Knudsen",
            "Simon Forsberg",
            "Stephen Bullen",
            "",
            "All contributors to our GitHub repository",
            "All our stargazers, likers & followers, for the warm fuzzies",
            "...and anyone reading this!"});
            this.SpecialThanksList.Location = new System.Drawing.Point(249, 202);
            this.SpecialThanksList.Margin = new System.Windows.Forms.Padding(4);
            this.SpecialThanksList.Name = "SpecialThanksList";
            this.SpecialThanksList.SelectionMode = System.Windows.Forms.SelectionMode.None;
            this.SpecialThanksList.Size = new System.Drawing.Size(400, 128);
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
            this.OfficialWebsiteLinkLabel.Location = new System.Drawing.Point(-1, 154);
            this.OfficialWebsiteLinkLabel.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.OfficialWebsiteLinkLabel.Name = "OfficialWebsiteLinkLabel";
            this.OfficialWebsiteLinkLabel.Size = new System.Drawing.Size(178, 20);
            this.OfficialWebsiteLinkLabel.TabIndex = 2;
            this.OfficialWebsiteLinkLabel.TabStop = true;
            this.OfficialWebsiteLinkLabel.Text = "rubberduck-vba.com";
            this.OfficialWebsiteLinkLabel.TextAlign = System.Drawing.ContentAlignment.TopCenter;
            // 
            // TwitterIcon
            // 
            this.TwitterIcon.Cursor = System.Windows.Forms.Cursors.Hand;
            this.TwitterIcon.Image = global::Rubberduck.Properties.Resources.twitter_circle_black_512;
            this.TwitterIcon.Location = new System.Drawing.Point(71, 327);
            this.TwitterIcon.Margin = new System.Windows.Forms.Padding(4);
            this.TwitterIcon.Name = "TwitterIcon";
            this.TwitterIcon.Size = new System.Drawing.Size(43, 39);
            this.TwitterIcon.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.TwitterIcon.TabIndex = 3;
            this.TwitterIcon.TabStop = false;
            // 
            // FacebookIcon
            // 
            this.FacebookIcon.Cursor = System.Windows.Forms.Cursors.Hand;
            this.FacebookIcon.Image = global::Rubberduck.Properties.Resources.facebook_circle_256;
            this.FacebookIcon.Location = new System.Drawing.Point(121, 329);
            this.FacebookIcon.Margin = new System.Windows.Forms.Padding(4);
            this.FacebookIcon.Name = "FacebookIcon";
            this.FacebookIcon.Size = new System.Drawing.Size(40, 37);
            this.FacebookIcon.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.FacebookIcon.TabIndex = 3;
            this.FacebookIcon.TabStop = false;
            // 
            // GitHubIcon
            // 
            this.GitHubIcon.Cursor = System.Windows.Forms.Cursors.Hand;
            this.GitHubIcon.Image = global::Rubberduck.Properties.Resources.github_circle_black_128;
            this.GitHubIcon.Location = new System.Drawing.Point(20, 327);
            this.GitHubIcon.Margin = new System.Windows.Forms.Padding(4);
            this.GitHubIcon.Name = "GitHubIcon";
            this.GitHubIcon.Size = new System.Drawing.Size(43, 39);
            this.GitHubIcon.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.GitHubIcon.TabIndex = 3;
            this.GitHubIcon.TabStop = false;
            // 
            // GooglePlusIcon
            // 
            this.GooglePlusIcon.Cursor = System.Windows.Forms.Cursors.Hand;
            this.GooglePlusIcon.Image = global::Rubberduck.Properties.Resources.google_circle_512;
            this.GooglePlusIcon.Location = new System.Drawing.Point(172, 329);
            this.GooglePlusIcon.Margin = new System.Windows.Forms.Padding(4);
            this.GooglePlusIcon.Name = "GooglePlusIcon";
            this.GooglePlusIcon.Size = new System.Drawing.Size(40, 37);
            this.GooglePlusIcon.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.GooglePlusIcon.TabIndex = 3;
            this.GooglePlusIcon.TabStop = false;
            // 
            // CloseButton
            // 
            this.CloseButton.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.CloseButton.Location = new System.Drawing.Point(61, 238);
            this.CloseButton.Margin = new System.Windows.Forms.Padding(4);
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
            this.CopyrightLabel.Location = new System.Drawing.Point(241, 347);
            this.CopyrightLabel.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.CopyrightLabel.Name = "CopyrightLabel";
            this.CopyrightLabel.Size = new System.Drawing.Size(328, 17);
            this.CopyrightLabel.TabIndex = 5;
            this.CopyrightLabel.Text = "© Copyright 2014-2015 Mathieu Guindon & Christopher McClellan";
            this.CopyrightLabel.UseMnemonic = false;
            // 
            // AttributionsLabel
            // 
            this.AttributionsLabel.AutoSize = true;
            this.AttributionsLabel.BackColor = System.Drawing.Color.Transparent;
            this.AttributionsLabel.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.AttributionsLabel.Location = new System.Drawing.Point(245, 66);
            this.AttributionsLabel.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.AttributionsLabel.Name = "AttributionsLabel";
            this.AttributionsLabel.Size = new System.Drawing.Size(106, 20);
            this.AttributionsLabel.TabIndex = 6;
            this.AttributionsLabel.Text = "Attributions";
            // 
            // SpecialThanksLabel
            // 
            this.SpecialThanksLabel.AutoSize = true;
            this.SpecialThanksLabel.BackColor = System.Drawing.Color.Transparent;
            this.SpecialThanksLabel.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.SpecialThanksLabel.Location = new System.Drawing.Point(245, 177);
            this.SpecialThanksLabel.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.SpecialThanksLabel.Name = "SpecialThanksLabel";
            this.SpecialThanksLabel.Size = new System.Drawing.Size(137, 20);
            this.SpecialThanksLabel.TabIndex = 6;
            this.SpecialThanksLabel.Text = "Special Thanks";
            // 
            // AttributionsList
            // 
            this.AttributionsList.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.AttributionsList.ForeColor = System.Drawing.Color.DimGray;
            this.AttributionsList.FormattingEnabled = true;
            this.AttributionsList.ItemHeight = 16;
            this.AttributionsList.Items.AddRange(new object[] {
            "Parsing powered by ANTLR",
            "GitHub integration powered by LibGit2Sharp",
            "Fugue icons by Yusuke Kamiyamane"});
            this.AttributionsList.Location = new System.Drawing.Point(249, 91);
            this.AttributionsList.Margin = new System.Windows.Forms.Padding(4);
            this.AttributionsList.Name = "AttributionsList";
            this.AttributionsList.SelectionMode = System.Windows.Forms.SelectionMode.None;
            this.AttributionsList.Size = new System.Drawing.Size(384, 64);
            this.AttributionsList.TabIndex = 1;
            // 
            // _AboutWindow
            // 
            this.AcceptButton = this.CloseButton;
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackgroundImage = global::Rubberduck.Properties.Resources.RD_AboutWindow;
            this.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None;
            this.CancelButton = this.CloseButton;
            this.ClientSize = new System.Drawing.Size(665, 382);
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
            this.Margin = new System.Windows.Forms.Padding(4);
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "_AboutWindow";
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