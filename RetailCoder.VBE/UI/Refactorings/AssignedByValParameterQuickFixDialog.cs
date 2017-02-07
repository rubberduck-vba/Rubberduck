using System;
using System.Linq;
using System.Windows.Forms;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Refactorings.Rename;

namespace Rubberduck.UI.Refactorings
{
    public partial class AssignedByValParameterQuickFixDialog : Form, IDialogView
    {
        private const string _INVALID_ENTRY = "Invalid Name:";
        private const string _QUESTIONABLE_ENTRY = "Note: Inspections for Maintainability and Readability will recommend changing a name like";

        private string[] _moduleLines;
        public AssignedByValParameterQuickFixDialog(string[] moduleLines)
        {
            _moduleLines = moduleLines;
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
                ValidateNewName();
            }
        }

        private void ValidateNewName()
        {

            bool isValidName = true;
            bool isMeaningfulName = true;
          
            if (IsEmpty())
            {
                isValidName = false;
            }
            else
            {
                isValidName = IsValidName(NewName);
                if (isValidName)
                {
                    isMeaningfulName = IsMeaningfulName(NewName);
                    if (!isMeaningfulName)
                    {
                        FeedbackLabel.Text = _QUESTIONABLE_ENTRY + " " + NewNameInQuotes();
                    }
                }
            }
            OkButton.Visible = isValidName;
            OkButton.Enabled = isValidName;
            InvalidNameValidationIcon.Visible = !isValidName;
            FeedbackLabel.Visible = !(isValidName  && isMeaningfulName);
        }
        private bool IsValidName(string identifier)
        {
            var tokenValues = typeof(Tokens).GetFields().Select(item => item.GetValue(null)).Cast<string>().Select(item => item);
            return !IsSameName()
                               && !FirstLetterIsDigit()
                               && !IsReservedToken(tokenValues)
                               && !UsesSpecialCharacters()
                               && !NewNameAlreadyUsed();
        }
        private bool IsMeaningfulName(string identifier)
        {
            return HasVowels()
                    && !NameIsASingleRepeatedLetter()
                    && !(NewName.Length < 3);
        }
        private bool IsEmpty()
        {
            if(NewName.Equals(string.Empty))
            {
                FeedbackLabel.Text = string.Empty;
            }
            return NewName.Equals(string.Empty);
        }
        private bool IsSameName()
        {
            if( NewName == Target.IdentifierName)
            {
                FeedbackLabel.Text = _INVALID_ENTRY + " " + NewNameInQuotes() + " is the ByVal parameter name";
                return true;
            }
            return false;
        }
        private bool FirstLetterIsDigit()
        {
            if (!char.IsLetter(NewName.FirstOrDefault()))
            {
                if (!NewName.Equals(string.Empty))
                {
                    FeedbackLabel.Text = _INVALID_ENTRY + "VBA variable names must start with a letter";
                }
                return true;
            }
            return false;
        }
        private bool IsReservedToken(System.Collections.Generic.IEnumerable<string> tokenValues)
        {
            if (tokenValues.Contains(NewName, StringComparer.InvariantCultureIgnoreCase))
            {
                FeedbackLabel.Text = _INVALID_ENTRY + " " + NewNameInQuotes() + " is a reserved VBA Word";
                return true;
            }
            return false;
        }
        private bool UsesSpecialCharacters()
        {
            if (NewName.Any(c => !char.IsLetterOrDigit(c) && c != '_'))
            {
                FeedbackLabel.Text = _INVALID_ENTRY + " " + "The variable name cannot include special character(s) except for '_'";
                return true;
            }
            return false;
        }
        private bool NewNameAlreadyUsed()
        {
            for(int idx = 0; idx < _moduleLines.Count();idx++)
            {
                string[] splitLine = _moduleLines[idx].Split(new char[] { ' ', ',' });
                if (splitLine.Contains(Tokens.Dim) && splitLine.Contains(NewName))
                {
                    FeedbackLabel.Text = _INVALID_ENTRY + NewNameInQuotes() + " is alread used in this code block";
                    return true;
                }
            }
           return false;
        }
        private bool HasVowels()
        {
            const string vowels = "aeiouyàâäéèêëïîöôùûü";
            return NewName.Any(character => vowels.Any(vowel =>
                   string.Compare(vowel.ToString(), character.ToString(), StringComparison.OrdinalIgnoreCase) == 0));
        }
        private bool NameIsASingleRepeatedLetter()
        {
            string firstLetter = NewName.First().ToString();
            return NewName.All(a => string.Compare(a.ToString(), firstLetter,
                StringComparison.OrdinalIgnoreCase) == 0);
        }
        private string NewNameInQuotes()
        {
            return "'" + NewName + "'";
        }
    }
}
