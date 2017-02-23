using System;
using System.Windows.Forms;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Inspections;
using Rubberduck.VBEditor;
using System.Linq;

namespace Rubberduck.UI.Refactorings
{
    public partial class AssignedByValParameterQuickFixDialog : Form, IDialogView
    {
        private string[] _identifierNamesAlreadyDeclared;

        public AssignedByValParameterQuickFixDialog(Declaration target, QualifiedSelection selection)
        {
            InitializeComponent();
            InitializeCaptions();
            Target = target;
            _identifierNamesAlreadyDeclared = Enumerable.Empty<string>().ToArray();
        }

        private void InitializeCaptions()
        {
            Text = RubberduckUI.AssignedByValParamQFixDialog_Caption;
            OkButton.Text = RubberduckUI.OK;
            CancelDialogButton.Text = RubberduckUI.CancelButtonText;
            TitleLabel.Text = RubberduckUI.AssignedByValParamQFixDialog_TitleText;
            InstructionsLabel.Text = RubberduckUI.AssignedByValParamQFixDialog_InstructionsLabelText;
            NameLabel.Text = RubberduckUI.NameLabelText;
        }

        private void NewNameBox_TextChanged(object sender, EventArgs e)
        {
            NewName = NewNameBox.Text;
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
                SetInstructionLableText();
            }
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

        private void SetInstructionLableText()
        {
            var declarationType =
                RubberduckUI.ResourceManager.GetString("DeclarationType_" + _target.DeclarationType, Settings.Settings.Culture);
            InstructionsLabel.Text = string.Format(RubberduckUI.AssignedByValParamQFixDialog_InstructionsLabelText, declarationType,
                _target.IdentifierName);
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
            return NewName.Equals(Target.IdentifierName, StringComparison.OrdinalIgnoreCase);
        }

        private bool NewNameAlreadyUsed()
        {
            return _identifierNamesAlreadyDeclared.Contains(NewName);
        }
    }
}
