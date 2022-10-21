using System.Drawing;
using System.Windows.Forms;
using Rubberduck.VersionCheck;

namespace Rubberduck.UI
{
    public partial class Splash2021 : Form
    {
        public Splash2021()
        {
            InitializeComponent();
        }

        public Splash2021(IVersionCheckService versionCheck) : this()
        {
            VersionLabel.Text = string.Format(Resources.RubberduckUI.Rubberduck_AboutBuild, versionCheck.VersionString);
            VersionLabel.Parent = pictureBox1;
            VersionLabel.BackColor = Color.Transparent;
        }
    }
}
