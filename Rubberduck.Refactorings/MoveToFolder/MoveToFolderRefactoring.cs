using System.Collections.Generic;
using System.Linq;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.Exceptions;
using Rubberduck.Refactorings.Exceptions.MoveToFolder;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.Utility;

namespace Rubberduck.Refactorings.MoveToFolder
{
    public class MoveToFolderRefactoring : InteractiveRefactoringBase<MoveMultipleToFolderModel>
    {
        private readonly IRefactoringAction<MoveMultipleToFolderModel> _moveToFolderAction;
        private readonly ISelectedDeclarationProvider _selectedDeclarationProvider;
        private readonly RubberduckParserState _state;

        public MoveToFolderRefactoring(
            MoveMultipleToFolderRefactoringAction moveToFolderAction,
            ISelectedDeclarationProvider selectedDeclarationProvider,
            ISelectionProvider selectionProvider, 
            RefactoringUserInteraction<IMoveMultipleToFolderPresenter, MoveMultipleToFolderModel> userInteraction,
            RubberduckParserState state) 
            : base(selectionProvider, userInteraction)
        {
            _moveToFolderAction = moveToFolderAction;
            _selectedDeclarationProvider = selectedDeclarationProvider;
            _state = state;
        }

        protected override Declaration FindTargetDeclaration(QualifiedSelection targetSelection)
        {
            return _selectedDeclarationProvider.SelectedModule(targetSelection);
        }

        protected override MoveMultipleToFolderModel InitializeModel(Declaration target)
        {
            if (target == null)
            {
                throw new TargetDeclarationIsNullException();
            }

            if (!(target is ModuleDeclaration targetModule))
            {
                throw new InvalidDeclarationTypeException(target);
            }
            
            var targets = new List<ModuleDeclaration>{ targetModule };
            var targetFolder = targetModule.CustomFolder;
            return new MoveMultipleToFolderModel(targets, targetFolder);
        }

        protected override void RefactorImpl(MoveMultipleToFolderModel model)
        {
            ValidateModel(model);
            _moveToFolderAction.Refactor(model);
        }

        private void ValidateModel(MoveMultipleToFolderModel model)
        {
            if (string.IsNullOrEmpty(model.TargetFolder))
            {
                throw new NoTargetFolderException();
            }

            var firstStaleAffectedModules = model.Targets
                .FirstOrDefault(module => _state.IsNewOrModified(module.QualifiedModuleName));
            if (firstStaleAffectedModules != null)
            {
                throw new AffectedModuleIsStaleException(firstStaleAffectedModules.QualifiedModuleName);
            }
        }
    }
}