using System;
using System.Collections.Generic;
using System.Windows.Forms;
using System.Linq;
using Rubberduck.Common;

namespace Rubberduck.UI.Refactorings
{
    public partial class AssignedByValParameterQuickFixDialog : Form, IAssignedByValParameterQuickFixDialog
    {
        private readonly IEnumerable<string> _forbiddenNames;

        public AssignedByValParameterQuickFixDialog(string identifier, string identifierType, IEnumerable<string> forbiddenNames)
        {
            InitializeComponent();
            InitializeCaptions(identifier, identifierType);
            _forbiddenNames = forbiddenNames;
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
            if (_forbiddenNames.Any(name => name.Equals(NewName, StringComparison.OrdinalIgnoreCase)))
            {
                return string.Format(RubberduckUI.AssignedByValDialog_NewNameAlreadyUsedFormat, NewName);
            }
            if (VariableNameValidator.StartsWithDigit(NewName))
            {
                return RubberduckUI.AssignedByValDialog_DoesNotStartWithLetter;
            }
            if (VariableNameValidator.HasSpecialCharacters(NewName))
            {
                return RubberduckUI.AssignedByValDialog_InvalidCharacters;
            }
            if (VariableNameValidator.IsReservedIdentifier(NewName))
            {
                return string.Format(RubberduckUI.AssignedByValDialog_ReservedKeywordFormat, NewName);
            }
            if (!VariableNameValidator.IsMeaningfulName(NewName))
            {
                return string.Format(RubberduckUI.AssignedByValDialog_QuestionableEntryFormat, NewName);
            }
            return string.Empty;
        }

        private void SetControlsProperties()
        {
            var isValid = VariableNameValidator.IsValidName(NewName) && !_forbiddenNames.Any(name => name.Equals(NewName, StringComparison.OrdinalIgnoreCase));
            OkButton.Visible = isValid;
            OkButton.Enabled = isValid;
            InvalidNameValidationIcon.Visible = !isValid;
        }
    }
}
