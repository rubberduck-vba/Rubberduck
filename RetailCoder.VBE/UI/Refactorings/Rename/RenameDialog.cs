using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
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
                NewName = value.IdentifierName;
                InstructionsLabel.Text = string.Format(RubberduckUI.RenameDialog_InstructionsLabelText,
                    value.DeclarationType.ToString().ToLower(), value.IdentifierName);
            }
        }

        public string NewName
        {
            get { return NewNameBox.Text; }
            set { NewNameBox.Text = value; }
        }
    }
}
