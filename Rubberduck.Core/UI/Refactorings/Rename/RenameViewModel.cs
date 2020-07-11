using System;
using System.Globalization;
using System.Linq;
using Rubberduck.Interaction;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.Rename;
using Rubberduck.Resources;
using Rubberduck.Refactorings.Common;

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
                ValidateName();
                OnPropertyChanged();
                OnPropertyChanged(nameof(IsValidName));
            }
        }
        
        private void ValidateName()
        {
            if (Target == null)
            {
                return;
            }

            var errors = VBAIdentifierValidator.SatisfiedInvalidIdentifierCriteria(NewName, Target.DeclarationType).ToList();

            var originalName = Model.Target.IdentifierName;
            if (!originalName.Equals(NewName)
                && originalName.Equals(NewName, StringComparison.InvariantCultureIgnoreCase))
            {
                errors.Add(RubberduckUI.RenameDialog_OnlyCasingDifferent);
            }

            if (errors.Any())
            {
                SetErrors(nameof(NewName), errors);
            }
            else
            {
                ClearErrors();
            }
        }

        public bool IsValidName => !HasErrors;

        protected override void DialogOk()
        {
            if (Target == null)
            {
                base.DialogCancel();
            }
            else
            {
                var firstConflictingDeclaration = _declarationFinderProvider.DeclarationFinder
                    .FindNewDeclarationNameConflicts(NewName, Model.Target)
                    .FirstOrDefault();

                if (firstConflictingDeclaration != null
                    && !UserConfirmsToProceedWithConflictingName(Model.NewName, Model.Target, firstConflictingDeclaration))
                {
                    base.DialogCancel();
                }
                else
                {
                    base.DialogOk();
                }
            }
        }

        private bool UserConfirmsToProceedWithConflictingName(string newName, Declaration target, Declaration conflictingDeclaration)
        {
            var message = string.Format(RubberduckUI.RenameDialog_ConflictingNames, newName, target.IdentifierName, conflictingDeclaration.QualifiedName.ToString());
            return _messageBox?.ConfirmYesNo(message, RubberduckUI.RenameDialog_Caption) ?? false;
        }
    }
}
