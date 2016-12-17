using System.Windows.Forms;

namespace Rubberduck.UI
{
    public partial class Splash : Form
    {
        public Splash()
        {
            InitializeComponent();
        }

        public string Version
        {
            get { return VersionLabel.Text; }
            set { VersionLabel.Text = value; }
        }
    }
}
