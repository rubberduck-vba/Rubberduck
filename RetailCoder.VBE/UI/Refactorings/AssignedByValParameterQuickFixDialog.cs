using System;
using System.Windows.Forms;
using Rubberduck.Inspections;
using System.Linq;

namespace Rubberduck.UI.Refactorings
{
    public partial class AssignedByValParameterQuickFixDialog : Form, IAssignedByValParameterQuickFixDialog
    {
        private string[] _identifierNamesAlreadyDeclared;
        private string _identifierName;

        internal AssignedByValParameterQuickFixDialog(string identifierName, string declarationType)
        {
            InitializeComponent();
            InitializeCaptions(identifierName, declarationType);
           _identifierName = identifierName;
            _identifierNamesAlreadyDeclared = Enumerable.Empty<string>().ToArray();
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
        public string[] IdentifierNamesAlreadyDeclared
        {
            get { return _identifierNamesAlreadyDeclared; }
            set { _identifierNamesAlreadyDeclared = value; }
        }

        private string GetVariableNameFeedback()
        {
            var validator = new VariableNameValidator(NewName);

            if (UserInputIsBlank())
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
            if (IsByValIdentifier())
            {
                return string.Format(RubberduckUI.AssignedByValDialog_IsByValIdentifierFormat, NewName);
            }
            if (NewNameAlreadyUsed())
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
            var userInputIsValid = validator.IsValidName() && !NewNameAlreadyUsed();
            OkButton.Visible = userInputIsValid;
            OkButton.Enabled = userInputIsValid;
            InvalidNameValidationIcon.Visible = !userInputIsValid;
        }

        private bool UserInputIsBlank()
        {
            return NewName.Equals(string.Empty);
        }

        private bool IsByValIdentifier()
        {
            return NewName.Equals(_identifierName, StringComparison.OrdinalIgnoreCase);
        }

        private bool NewNameAlreadyUsed()
        {
            //Comparison needs to be case-insensitive, or VBE will often change an existing
            //same-spelling local variable's casing to conform with the NewName
            return _identifierNamesAlreadyDeclared.Any(n => n.Equals(NewName, StringComparison.OrdinalIgnoreCase));
        }
    }
}
