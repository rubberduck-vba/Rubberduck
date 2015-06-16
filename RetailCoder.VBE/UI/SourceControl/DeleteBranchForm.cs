using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Rubberduck.UI.SourceControl
{
    public partial class DeleteBranchForm : Form, IDeleteBranchView
    {
        public DeleteBranchForm()
        {
            InitializeComponent();

            Branches = new List<string>();
        }

        public bool OkButtonEnabled
        {
            get { return this.OkButton.Enabled; }
            set { this.OkButton.Enabled = value; }
        }

        private IList<string> _branches;
        public IList<string> Branches
        {
            get { return _branches; }
            set
            {
                _branches = value;
                BranchList.DataSource = Branches;
                BranchList.Refresh();
           }
        }

        public event EventHandler<BranchDeleteArgs> Confirm;
        private void OkButton_Click(object sender, EventArgs e)
        {
            var handler = Confirm;
            if (handler != null)
            {
                handler(this, new BranchDeleteArgs(this.BranchList.SelectedItem.ToString()));
            }
        }

        public event EventHandler<EventArgs> Cancel;
        private void CancelButton_Click(object sender, EventArgs e)
        {
            var handler = Cancel;
            if (handler != null)
            {
                handler(this, e);
            }
        }
    }
}
