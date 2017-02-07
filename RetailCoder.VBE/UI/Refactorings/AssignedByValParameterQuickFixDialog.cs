using System;
using System.Linq;
using System.Windows.Forms;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Inspections.QuickFixes;
using Rubberduck.Inspections;

namespace Rubberduck.UI.Refactorings
{
    public partial class AssignedByValParameterQuickFixDialog : Form, IDialogView
    {
        private const string _INVALID_ENTRY_PROLOGUE = "Invalid Name:";

        private string[] _moduleLines;
        private bool _userInputIsValid;

        public AssignedByValParameterQuickFixDialog(string[] moduleLines)
        {
            _moduleLines = moduleLines;
            _userInputIsValid = false;
            InitializeComponent();
            InitializeCaptions();
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

                var declarationType =
                    RubberduckUI.ResourceManager.GetString("DeclarationType_" + _target.DeclarationType, UI.Settings.Settings.Culture);
                InstructionsLabel.Text = string.Format(RubberduckUI.AssignedByValParamQFixDialog_InstructionsLabelText, declarationType,
                    _target.IdentifierName);
            }
        }

        public string NewName
        {
            get { return NewNameBox.Text; }
            set
            {
                NewNameBox.Text = value;
                FeedbackLabel.Text = !value.Equals(string.Empty) ? GetVariableNameFeedback() : string.Empty;
                _userInputIsValid = !FeedbackLabel.Text.StartsWith(_INVALID_ENTRY_PROLOGUE)
                                    && !value.Equals(string.Empty);
                SetControlsProperties();
            }
        }
        private string GetVariableNameFeedback()
        {
            var validator = new VariableNameValidator(NewName);
            if (UserInputIsBlank()) { return string.Empty; }
            if (validator.StartsWithNumber) { return InvalidEntryMsg("VBA variable names must start with a letter"); }
            if (validator.ContainsSpecialCharacters) { return InvalidEntryMsg("VBA variable names cannot include special character(s) except for '_'"); }
            if (validator.IsReservedName) { return InvalidEntryMsg(NewNameInQuotes() + " is a reserved VBA Word"); }
            if (NewNameAlreadyUsed()) { return InvalidEntryMsg(NewNameInQuotes() + " is already used in this code block"); }
            if (IsByValIdentifier()) { return InvalidEntryMsg(NewNameInQuotes() + " is the ByVal parameter name"); }
            if (!validator.IsMeaningfulName()) { return QuestionableEntryMsg(); }
            return string.Empty;
        }
        private void SetControlsProperties()
        {
            OkButton.Visible = _userInputIsValid;
            OkButton.Enabled = _userInputIsValid;
            InvalidNameValidationIcon.Visible = !_userInputIsValid;
        }
        private bool UserInputIsBlank()
        {
            return NewName.Equals(string.Empty);
        }
        private bool IsByValIdentifier()
        {
            return NewName.Equals(Target.IdentifierName,StringComparison.OrdinalIgnoreCase);
        }
        private bool NewNameAlreadyUsed()
        {
            for(int idx = 0; idx < _moduleLines.Count();idx++)
            {
                string[] splitLine = _moduleLines[idx].ToUpper().Split(new char[] { ' ', ',' });
                if( splitLine.Contains(Tokens.Dim.ToUpper()) && splitLine.Contains(NewName.ToUpper()))
                {
                    return true;
                }
            }
            return false;
        }
        private string NewNameInQuotes()
        {
            return "'" + NewName + "'";
        }
        private string InvalidEntryMsg(string message)
        {
            return _INVALID_ENTRY_PROLOGUE + " " + message;
        }
        private string QuestionableEntryMsg()
        {
            const string _QUESTIONABLE_ENTRY = "Note: A name like '{0}' will be"
                    + " identified as a 'Maintainability and Readability Issue'."
                    + "  Consider choosing a different name.";

            return string.Format(_QUESTIONABLE_ENTRY, NewName);
        }
    }
}
