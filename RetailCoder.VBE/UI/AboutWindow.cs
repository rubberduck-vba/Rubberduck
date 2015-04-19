using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace Rubberduck.UI
{
    [ComVisible(true)]
    [Guid(ClassId)]
    [ProgId(ProgId)]
    public partial class _AboutWindow : Form
    {
        public const string ClassId = "939CC8BB-A8CA-3BE6-89A3-5450949A6A43";
        public const string ProgId = "Rubberduck.UI._AboutWindow";

        private static readonly IDictionary<string, string> Links =
            new Dictionary<string, string>
            {
                {"Home","http://www.rubberduck-vba.com"},
                {"GitHub", "http://www.github.com/retailcoder/rubberduck"},
                {"Twitter","http://www.twitter.com/rubberduckvba"},
                {"Facebook", "http://www.facebook.com/rubberduckvba"},
                {"Google+", "http://plus.google.com/116859653258584466987"}
            };

        public _AboutWindow()
        {
            InitializeComponent();
            var assembly = Assembly.GetExecutingAssembly();
            var name = assembly.GetName();

            AppVersionLabel.Text = string.Format("Build {0} ({1})", name.Version, name.ProcessorArchitecture);
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
            VisitLink(Links["Home"]);
        }

        private void CloseButton_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void VisitLink(string url)
        {
            var info = new ProcessStartInfo(url);
            Process.Start(info);
        }
    }
}
