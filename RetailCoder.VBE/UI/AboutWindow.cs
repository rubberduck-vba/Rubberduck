using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Rubberduck
{
    public partial class AboutWindow : Form
    {
        public AboutWindow()
        {
            InitializeComponent();
            var assembly = System.Reflection.Assembly.GetExecutingAssembly();
            var name = assembly.GetName();

            titleLabel.Text = name.Name;
            versionLabel.Text = name.Version.ToString();

            repositoryLinkLabel.LinkClicked += repositoryLinkLabel_LinkClicked;
            retailcoderLinkLabel.LinkClicked += retailcoderLinkLabel_LinkClicked;
            ckuhn203LinkLabel.LinkClicked += ckuhn203LinkLabel_LinkClicked;

            fugueiconsLinkLabel.LinkClicked += fugueiconsLinkLabel_LinkClicked;
            codereviewLinkLabel.LinkClicked += codereviewLinkLabel_LinkClicked;
            mztoolsLinkLabel.LinkClicked += mztoolsLinkLabel_LinkClicked;

            CloseButton.Click += CloseButton_Click;
        }

        void CloseButton_Click(object sender, EventArgs e)
        {
            Close();
        }

        void mztoolsLinkLabel_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            VisitLink("http://www.mztools.com/articles/2006/mz2006007.aspx");
        }

        void codereviewLinkLabel_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            VisitLink("http://www.codereview.stackexchange.com");
        }

        void fugueiconsLinkLabel_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            VisitLink("http://p.yusukekamiyamane.com");
        }

        void ckuhn203LinkLabel_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            VisitLink("https://github.com/ckuhn203");
        }

        void retailcoderLinkLabel_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            VisitLink("https://github.com/retailcoder");
        }

        void repositoryLinkLabel_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            VisitLink("https://github.com/retailcoder/Rubberduck");
        }

        private void VisitLink(string url)
        {
            var info = new ProcessStartInfo(url);
            Process.Start(info);
        }
    }
}
