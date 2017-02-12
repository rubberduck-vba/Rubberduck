using System;
using System.Linq;
using System.Windows.Forms;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Inspections;

namespace Rubberduck.UI.Refactorings
{
    public partial class AssignedByValParameterQuickFixDialog : Form, IDialogView
    {
        private readonly string[] _moduleLines;
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
                    RubberduckUI.ResourceManager.GetString("DeclarationType_" + _target.DeclarationType, Settings.Settings.Culture);
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
                SetControlsProperties();
            }
        }

        private string GetVariableNameFeedback()
        {
            var validator = new VariableNameValidator(NewName);
            _userInputIsValid = validator.IsValidName() && !NewNameAlreadyUsed();

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
            if (NewNameAlreadyUsed())
            {
                return string.Format(RubberduckUI.AssignedByValDialog_NewNameAlreadyUsedFormat, NewName);
            }
            if (IsByValIdentifier())
            {
                return string.Format(RubberduckUI.AssignedByValDialog_IsByValIdentifierFormat, NewName);
            }
            if (!validator.IsMeaningfulName())
            {
                return string.Format(RubberduckUI.AssignedByValDialog_QuestionableEntryFormat, NewName);
            }
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
            var validator = new VariableNameValidator(NewName);
            return _moduleLines.Any(codeLine => validator.IsReferencedIn(codeLine));
        }
    }
}
