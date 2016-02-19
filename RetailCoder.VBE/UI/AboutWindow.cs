using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace Rubberduck.UI
{
    [ComVisible(false)]
    // ReSharper disable once InconsistentNaming
    public sealed partial class AboutWindow : Form
    {
        private static readonly IDictionary<string, string> Links =
            new Dictionary<string, string>
            {
                {RubberduckUI.Home,"http://www.rubberduck-vba.com"},
                {"GitHub", "http://www.github.com/retailcoder/rubberduck"},
                {"Twitter","http://www.twitter.com/rubberduckvba"},
                {"Facebook", "http://www.facebook.com/rubberduckvba"},
                {"Google+", "http://plus.google.com/116859653258584466987"}
            };

        public AboutWindow()
        {
            InitializeComponent();
            var assembly = Assembly.GetExecutingAssembly();
            var name = assembly.GetName();

            Text = RubberduckUI.About_Caption;
            AttributionsLabel.Text = RubberduckUI.AboutWindow_AttributionsLabel;
            SpecialThanksLabel.Text = RubberduckUI.About_SpecialThanks;
            CopyrightLabel.Text = RubberduckUI.About_Copyright;

            AppVersionLabel.Text = string.Format(RubberduckUI.Rubberduck_AboutBuild, name.Version, name.ProcessorArchitecture);
            CloseButton.Click += CloseButton_Click;

            OfficialWebsiteLinkLabel.LinkClicked += OfficialWebsiteLinkLabel_LinkClicked;
            GitHubIcon.Click += GitHubIcon_Click;
            TwitterIcon.Click += TwitterIcon_Click;
            FacebookIcon.Click += FacebookIcon_Click;
            GooglePlusIcon.Click += GooglePlusIcon_Click;
        }

        private void GooglePlusIcon_Click(object sender, EventArgs e)
        {
            VisitLink(Links["Google+"]);
        }

        private void FacebookIcon_Click(object sender, EventArgs e)
        {
            VisitLink(Links["Facebook"]);
        }

        private void TwitterIcon_Click(object sender, EventArgs e)
        {
            VisitLink(Links["Twitter"]);
        }

        private void GitHubIcon_Click(object sender, EventArgs e)
        {
            VisitLink(Links["GitHub"]);
        }

        private void OfficialWebsiteLinkLabel_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            VisitLink(Links[RubberduckUI.Home]);
        }

        private void CloseButton_Click(object sender, EventArgs e)
        {
            Close();
        }

        private static void VisitLink(string url)
        {
            var info = new ProcessStartInfo(url);
            Process.Start(info);
        }
    }
}
