using System;
using System.Globalization;
using System.Linq;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.Rename;
using Rubberduck.Resources;

namespace Rubberduck.UI.Refactorings.Rename
{
    public class RenameViewModel : RefactoringViewModelBase<RenameModel>
    {
        public RubberduckParserState State { get; }

        public RenameViewModel(RubberduckParserState state, RenameModel model) : base(model)
        {
            State = state;
        }

        public Declaration Target
        {
            get => Model.Target;
            set
            {
                Model.Target = value;
                NewName = Model.Target.IdentifierName;

                OnPropertyChanged(nameof(Instructions));
            }
        }

        public string Instructions
        {
            get
            {
                if (Target == null)
                {
                    return string.Empty;
                }

                var declarationType = RubberduckUI.ResourceManager.GetString("DeclarationType_" + Target.DeclarationType, CultureInfo.CurrentUICulture);
                return string.Format(RubberduckUI.RenameDialog_InstructionsLabelText, declarationType, Target.IdentifierName);
            }
        }

        public string NewName
        {
            get => Model.NewName;
            set
            {
                Model.NewName = value;
                OnPropertyChanged();
                OnPropertyChanged(nameof(IsValidName));
            }
        }
        
        public bool IsValidName
        {
            get
            {
                if (Target == null) { return false; }

                var tokenValues = typeof(Tokens).GetFields().Select(item => item.GetValue(null)).Cast<string>().Select(item => item);

                return !(NewName.Equals(Target.IdentifierName, StringComparison.InvariantCultureIgnoreCase)) &&
                       char.IsLetter(NewName.FirstOrDefault()) &&
                       !tokenValues.Contains(NewName, StringComparer.InvariantCultureIgnoreCase) &&
                       !NewName.Any(c => !char.IsLetterOrDigit(c) && c != '_') &&
                       NewName.Length <= (Target.DeclarationType.HasFlag(DeclarationType.Module) ? Declaration.MaxModuleNameLength : Declaration.MaxMemberNameLength);
            }
        }
    }
}
