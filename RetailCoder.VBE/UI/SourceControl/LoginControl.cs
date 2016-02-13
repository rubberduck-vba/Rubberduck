using System;
using System.Diagnostics.CodeAnalysis;
using System.Windows.Forms;

namespace Rubberduck.UI.SourceControl
{
    [ExcludeFromCodeCoverage]
    public partial class LoginControl : UserControl, ILoginView
    {
        public LoginControl()
        {
            InitializeComponent();
        }

        public string Username
        {
            get { return UsernameBox.Text; }
            set { UsernameBox.Text = value; }
        }

        public string Password
        {
            get { return PasswordBox.Text; }
            set { PasswordBox.Text = value; }
        }

        public event EventHandler Confirm;
        public event EventHandler Cancel;
        public event EventHandler<EventArgs> DismissSecondaryPanel;

        private void OkButton_Click(object sender, EventArgs e)
        {
            var handler = Confirm;
            if (handler != null)
            {
                handler(this, e);
            }

            var dismiss = DismissSecondaryPanel;
            if (dismiss != null)
            {
                dismiss(this, e);
            }
        }

        private void CancelButton_Click(object sender, EventArgs e)
        {
            var handler = Confirm;
            if (handler != null)
            {
                handler(this, e);
            }

            var dismiss = DismissSecondaryPanel;
            if (dismiss != null)
            {
                dismiss(this, e);
            }
        }
    }
}
