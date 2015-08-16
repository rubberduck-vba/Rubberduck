using System;
using System.Diagnostics.CodeAnalysis;
using System.Windows.Forms;

namespace Rubberduck.UI.SourceControl
{
    [ExcludeFromCodeCoverage]
    public partial class CreateBranchForm : Form, ICreateBranchView
    {
        public CreateBranchForm()
        {
            InitializeComponent();

            Text = RubberduckUI.SourceControl_CreateNewBranchCaption;
            TitleLabel.Text = RubberduckUI.SourceControl_CreateNewBranchTitle;
            InstructionsLabel.Text = RubberduckUI.SourceControl_CreateNewBranchInstructions;
            OkButton.Text = RubberduckUI.OK_AllCaps;
            OkButton.Click += OkButton_Click;
            CancelButton.Text = RubberduckUI.CancelButtonText;
            CancelButton.Click += CancelButton_Click;
        }

        public string UserInputText
        {
            get { return this.UserInputBox.Text; }
            set { this.UserInputBox.Text = value; }
        }

        public bool IsValidBranchName
        {
            get { return this.OkButton.Enabled; }
            set
            {
                this.OkButton.Enabled = value;
                this.InvalidNameValidationIcon.Visible = !value;
            }
        }

        public event EventHandler<BranchCreateArgs> Confirm;
        private void OkButton_Click(object sender, EventArgs e)
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
