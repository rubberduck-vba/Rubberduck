using System;
using System.Windows.Forms;
using Rubberduck.Parsing.Symbols;

namespace Rubberduck.UI.Refactorings.Rename
{
    public partial class RenameDialog : Form, IRenameView
    {
        public RenameDialog()
        {
            InitializeComponent();
            OkButton.Click += OkButtonClick;
        }

        private void OkButtonClick(object sender, EventArgs e)
        {
            OnOkButtonClicked();
        }

        public event EventHandler CancelButtonClicked;
        public void OnCancelButtonClicked()
        {
            Hide();
        }

        public event EventHandler OkButtonClicked;
        public void OnOkButtonClicked()
        {
            var handler = OkButtonClicked;
            if (handler != null)
            {
                handler(this, EventArgs.Empty);
            }
        }

        private Declaration _target;

        public Declaration Target
        {
            get { return _target; }
            set
            {
                _target = value;
                if (_target == null)
                {
                    return;
                }

                NewName = _target.IdentifierName;
                var declarationType = RubberduckUI.ResourceManager.GetString("DeclarationType_" + _target.DeclarationType);
                InstructionsLabel.Text = string.Format(RubberduckUI.RenameDialog_InstructionsLabelText, declarationType, _target.IdentifierName);
            }
        }

        public string NewName
        {
            get { return NewNameBox.Text; }
            set { NewNameBox.Text = value; }
        }
    }
}
