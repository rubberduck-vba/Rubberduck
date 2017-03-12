using System;
using System.Globalization;
using System.Linq;
using System.Windows.Forms;
using NLog;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.UI.Command;

namespace Rubberduck.UI.Refactorings.Rename
{
    public class RenameViewModel : ViewModelBase
    {
        public RubberduckParserState State { get; }
        public Declaration Target { get; }

        public RenameViewModel(RubberduckParserState state, Declaration target)
        {
            State = state;
            Target = target;

            NewName = Target.IdentifierName;

            OkButtonCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), _ => DialogOk());
            CancelButtonCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), _ => DialogCancel());
        }
        
        public string Instructions
        {
            get
            {
                var declarationType = RubberduckUI.ResourceManager.GetString("DeclarationType_" + Target.DeclarationType, CultureInfo.CurrentUICulture);
                return string.Format(RubberduckUI.RenameDialog_InstructionsLabelText, declarationType, Target.IdentifierName);
            }
        }

        private string _newName;
        public string NewName
        {
            get { return _newName; }
            set
            {
                _newName = value;
                OnPropertyChanged();
                OnPropertyChanged(nameof(IsValidName));
            }
        }
        
        public bool IsValidName
        {
            get
            {
                var tokenValues = typeof(Tokens).GetFields().Select(item => item.GetValue(null)).Cast<string>().Select(item => item);

                return NewName != Target.IdentifierName &&
                       char.IsLetter(NewName.FirstOrDefault()) &&
                       !tokenValues.Contains(NewName, StringComparer.InvariantCultureIgnoreCase) &&
                       !NewName.Any(c => !char.IsLetterOrDigit(c) && c != '_');
            }
        }

        public event EventHandler<DialogResult> OnWindowClosed;
        private void DialogCancel() => OnWindowClosed?.Invoke(this, DialogResult.Cancel);
        private void DialogOk() => OnWindowClosed?.Invoke(this, DialogResult.OK);
        
        public CommandBase OkButtonCommand { get; }
        public CommandBase CancelButtonCommand { get; }
    }
}
