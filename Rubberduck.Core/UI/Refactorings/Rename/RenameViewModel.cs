using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using Rubberduck.Interaction;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.Rename;
using Rubberduck.Resources;

namespace Rubberduck.UI.Refactorings.Rename
{
    public class RenameViewModel : RefactoringViewModelBase<RenameModel>
    {
        private readonly IDeclarationFinderProvider _declarationFinderProvider;
        private readonly IMessageBox _messageBox;

        public RubberduckParserState State { get; }

        public RenameViewModel(RubberduckParserState state, RenameModel model, IMessageBox messageBox) 
            : base(model)
        {
            State = state;
            _declarationFinderProvider = state;
            _messageBox = messageBox;
        }

        public Declaration Target => Model.Target;

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
                if (Target == null)
                {
                    return false;
                }

                var tokenValues = typeof(Tokens).GetFields().Select(item => item.GetValue(null)).Cast<string>().Select(item => item);

                return !(NewName.Equals(Target.IdentifierName, StringComparison.InvariantCultureIgnoreCase)) &&
                       char.IsLetter(NewName.FirstOrDefault()) &&
                       !tokenValues.Contains(NewName, StringComparer.InvariantCultureIgnoreCase) &&
                       !NewName.Any(c => !char.IsLetterOrDigit(c) && c != '_') &&
                       NewName.Length <= (Target.DeclarationType.HasFlag(DeclarationType.Module) ? Declaration.MaxModuleNameLength : Declaration.MaxMemberNameLength);
            }
        }

        protected override void DialogOk()
        {
            if (Target == null
                || (DeclarationsWithConflictingName(Model.NewName, Model.Target).Any()
                    && !UserConfirmsToProceedWithConflictingName(Model.NewName, Model.Target)))
            {
                base.DialogCancel();
            }
            else
            {
                base.DialogOk();
            }
        }

        private IEnumerable<Declaration> DeclarationsWithConflictingName(string newName, Declaration target)
        {
            return _declarationFinderProvider.DeclarationFinder.FindNewDeclarationNameConflicts(newName, target);
        }

        private bool UserConfirmsToProceedWithConflictingName(string newName, Declaration target)
        {
            var message = string.Format(RubberduckUI.RenameDialog_ConflictingNames, newName, target.IdentifierName);
            return _messageBox?.ConfirmYesNo(message, RubberduckUI.RenameDialog_Caption) ?? false;
        }
    }
}
