using System;
using System.Collections.Generic;
using System.Windows.Forms;
using Rubberduck.Inspections;
using System.Linq;

namespace Rubberduck.UI.Refactorings
{
    public partial class AssignedByValParameterQuickFixDialog : Form, IAssignedByValParameterQuickFixDialog
    {
        private readonly string _identifierName;
        private readonly IEnumerable<string> _forbiddenNames;

        internal AssignedByValParameterQuickFixDialog(string identifierName, string declarationType, IEnumerable<string> forbiddenNames)
        {
            InitializeComponent();
            InitializeCaptions(identifierName, declarationType);
           _identifierName = identifierName;
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
            var validator = new VariableNameValidator(NewName);

            if (string.IsNullOrEmpty(NewName))
            {
                return string.Empty;
            }
            if (validator.StartsWithNumber)
            {
                return RubberduckUI.AssignedByValDialog_DoesNotStartWithLetter;
            }
            if (validator.ContainsSpecialCharacters)
            {
                return RubberduckUI.AssignedByValDialog_InvalidCharacters;
            }
            if (validator.IsReservedName)
            {
                return string.Format(RubberduckUI.AssignedByValDialog_ReservedKeywordFormat, NewName);
            }
            if (NewName.Equals(_identifierName, StringComparison.OrdinalIgnoreCase))
            {
                return string.Format(RubberduckUI.AssignedByValDialog_IsByValIdentifierFormat, NewName);
            }
            if (_forbiddenNames.Any(name => name.Equals(NewName, StringComparison.OrdinalIgnoreCase)))
            {
                return string.Format(RubberduckUI.AssignedByValDialog_NewNameAlreadyUsedFormat, NewName);
            }
            if (!validator.IsMeaningfulName())
            {
                return string.Format(RubberduckUI.AssignedByValDialog_QuestionableEntryFormat, NewName);
            }
            return string.Empty;
        }

        private void SetControlsProperties()
        {
            var validator = new VariableNameValidator(NewName);
            var isValid = validator.IsValidName() && !_forbiddenNames.Any(name => name.Equals(NewName, StringComparison.OrdinalIgnoreCase));
            OkButton.Visible = isValid;
            OkButton.Enabled = isValid;
            InvalidNameValidationIcon.Visible = !isValid;
        }
    }
}
