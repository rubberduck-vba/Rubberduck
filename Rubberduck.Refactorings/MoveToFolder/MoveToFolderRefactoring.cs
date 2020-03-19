using System.Collections.Generic;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.UIContext;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.Exceptions;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.Utility;

namespace Rubberduck.Refactorings.MoveToFolder
{
    public class MoveToFolderRefactoring : InteractiveRefactoringBase<IMoveMultipleToFolderPresenter, MoveMultipleToFolderModel>
    {
        private readonly IRefactoringAction<MoveMultipleToFolderModel> _moveToFolderAction;
        private readonly ISelectedDeclarationProvider _selectedDeclarationProvider;

        public MoveToFolderRefactoring(
            MoveMultipleToFolderRefactoringAction moveToFolderAction,
            ISelectedDeclarationProvider selectedDeclarationProvider,
            ISelectionProvider selectionProvider, 
            IRefactoringPresenterFactory factory, 
            IUiDispatcher uiDispatcher) 
            : base(selectionProvider, factory, uiDispatcher)
        {
            _moveToFolderAction = moveToFolderAction;
            _selectedDeclarationProvider = selectedDeclarationProvider;
        }

        protected override Declaration FindTargetDeclaration(QualifiedSelection targetSelection)
        {
            return _selectedDeclarationProvider.SelectedModule(targetSelection);
        }

        protected override MoveMultipleToFolderModel InitializeModel(Declaration target)
        {
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
            _moveToFolderAction.Refactor(model);
        }
    }
}