using System;
using System.Linq;
using System.Windows.Forms;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Refactorings.Rename;

namespace Rubberduck.UI.Refactorings
{
    public partial class RenameDialog : Form, IRenameDialog
    {
        public RenameDialog()
        {
            InitializeComponent();
            InitializeCaptions();

            Shown += RenameDialog_Shown;
            NewNameBox.TextChanged += NewNameBox_TextChanged;
        }

        private void InitializeCaptions()
        {
            Text = RubberduckUI.RenameDialog_Caption;
            OkButton.Text = RubberduckUI.OK;
            CancelDialogButton.Text = RubberduckUI.CancelButtonText;
            TitleLabel.Text = RubberduckUI.RenameDialog_TitleText;
            InstructionsLabel.Text = RubberduckUI.RenameDialog_InstructionsLabelText;
            NameLabel.Text = RubberduckUI.NameLabelText;
        }

        private void NewNameBox_TextChanged(object sender, EventArgs e)
        {
            NewName = NewNameBox.Text;
        }

        private void RenameDialog_Shown(object sender, EventArgs e)
        {
            NewNameBox.SelectAll();
            NewNameBox.Focus();
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
                var declarationType =
                    RubberduckUI.ResourceManager.GetString("DeclarationType_" + _target.DeclarationType, RubberduckUI.Culture);
                InstructionsLabel.Text = string.Format(RubberduckUI.RenameDialog_InstructionsLabelText, declarationType,
                    _target.IdentifierName);
            }
        }

        public string NewName
        {
            get { return NewNameBox.Text; }
            set
            {
                NewNameBox.Text = value;
                ValidateNewName();
            }
        }

        private void ValidateNewName()
        {
            var tokenValues = typeof(Tokens).GetFields().Select(item => item.GetValue(null)).Cast<string>().Select(item => item);

            OkButton.Enabled = NewName != Target.IdentifierName
                               && char.IsLetter(NewName.FirstOrDefault())
                               && !tokenValues.Contains(NewName, StringComparer.InvariantCultureIgnoreCase)
                               && !NewName.Any(c => !char.IsLetterOrDigit(c) && c != '_');

            InvalidNameValidationIcon.Visible = !OkButton.Enabled;
        }
    }
}
