using System;
using System.Windows.Forms;
using Rubberduck.Common;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Refactorings.Common;
using Rubberduck.Resources;

namespace Rubberduck.UI.Refactorings
{
    public partial class AssignedByValParameterQuickFixDialog : Form, IAssignedByValParameterQuickFixDialog
    {
        private readonly Func<string, bool> _isConflictingName;

        public AssignedByValParameterQuickFixDialog(string identifier, string identifierType, Func<string, bool> nameCollisionChecker)
        {
            InitializeComponent();
            InitializeCaptions(identifier, identifierType);
            _isConflictingName = nameCollisionChecker;
        }

        private void InitializeCaptions(string identifierName, string targetDeclarationType)
        {
            Text = RubberduckUI.AssignedByValParamQFixDialog_Caption;
            OkButton.Text = RubberduckUI.OK;
            CancelDialogButton.Text = RubberduckUI.CancelButtonText;
            TitleLabel.Text = RubberduckUI.AssignedByValParamQFixDialog_TitleText;
            NameLabel.Text = RubberduckUI.NameLabelText;

            var declarationType =
            RubberduckUI.ResourceManager.GetString("DeclarationType_" + targetDeclarationType, Settings.Settings.Culture);
            InstructionsLabel.Text = string.Format(RubberduckUI.AssignedByValParamQFixDialog_InstructionsLabelText, declarationType,
                identifierName);
        }

        private void NewNameBox_TextChanged(object sender, EventArgs e)
        {
            NewName = NewNameBox.Text;
        }

        public string NewName
        {
            get { return NewNameBox.Text; }
            set
            {
                NewNameBox.Text = value;
                FeedbackLabel.Text = !value.Equals(string.Empty) ? GetVariableNameFeedback() : string.Empty;
                SetControlsProperties();
            }
        }

        private string GetVariableNameFeedback()
        {
            if (string.IsNullOrEmpty(NewName))
            {
                return string.Empty;
            }

            if (_isConflictingName(NewName))
            {
                return string.Format(RubberduckUI.AssignedByValDialog_NewNameAlreadyUsedFormat, NewName);
            }

            if (VBAIdentifierValidator.TryMatchInvalidIdentifierCriteria(NewName, DeclarationType.Variable, out var invalidMessage))
            {
                return invalidMessage;
            }

            if (VBAIdentifierValidator.TryMatchMeaninglessIdentifierCriteria(NewName, out var meaninglessNameMessage))
            {
                return string.Format(RubberduckUI.AssignedByValDialog_MeaninglessNameFormat, meaninglessNameMessage);
            }

            return string.Empty;
        }

        private void SetControlsProperties()
        {
            var isValid = VBAIdentifierValidator.IsValidIdentifier(NewName, DeclarationType.Variable);
            OkButton.Visible = isValid;
            OkButton.Enabled = isValid;
            InvalidNameValidationIcon.Visible = !isValid;
        }
    }
}
