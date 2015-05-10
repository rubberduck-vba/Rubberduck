using System;
using System.Windows.Forms;

namespace Rubberduck.UI.SourceControl
{
    public partial class CreateBranchForm : Form, ICreateBranchView
    {
        public CreateBranchForm()
        {
            InitializeComponent();
        }

        public string UserInputText
        {
            get { return this.UserInputBox.Text; }
            set { this.UserInputBox.Text = value; }
        }

        public bool OkayButtonEnabled
        {
            get { return this.OkayButton.Enabled; }
            set { this.OkayButton.Enabled = value; }
        }

        public event EventHandler<BranchCreateArgs> Confirm;
        private void Okay_Click(object sender, EventArgs e)
        {
            var handler = Confirm;
            if (handler != null)
            {
                handler(this, new BranchCreateArgs(this.UserInputText));
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

        public event EventHandler<EventArgs> UserInputTextChanged;
        private void UserInputBox_TextChanged(object sender, EventArgs e)
        {
            var handler = UserInputTextChanged;
            if (handler != null)
            {
                handler(this, e);
            }
            
        }
    }
}
