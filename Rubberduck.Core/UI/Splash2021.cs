using System.Drawing;
using System.Windows.Forms;

namespace Rubberduck.UI
{
    public partial class Splash2021 : Form
    {
        public Splash2021()
        {
            InitializeComponent();
        }

        public Splash2021(string versionString) : this()
        {
            VersionLabel.Text = string.Format(Resources.RubberduckUI.Rubberduck_AboutBuild, versionString);
            VersionLabel.Parent = pictureBox1;
            VersionLabel.BackColor = Color.Transparent;
        }
    }
}
