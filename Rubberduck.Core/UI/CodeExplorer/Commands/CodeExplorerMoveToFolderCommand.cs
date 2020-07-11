using System;
using System.Collections.Generic;
using Rubberduck.Navigation.CodeExplorer;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings;
using Rubberduck.Refactorings.MoveFolder;
using Rubberduck.Refactorings.MoveToFolder;
using Rubberduck.UI.CodeExplorer.Commands.Abstract;
using Rubberduck.VBEditor.Events;
using Rubberduck.UI.Command.Refactorings.Notifiers;

namespace Rubberduck.UI.CodeExplorer.Commands
{
    public sealed class CodeExplorerMoveToFolderCommand : CodeExplorerMoveToFolderCommandBase
    {
        private readonly IRefactoringUserInteraction<MoveMultipleFoldersModel> _moveFoldersInteraction;
        private readonly IRefactoringUserInteraction<MoveMultipleToFolderModel> _moveToFolderInteraction;

        public CodeExplorerMoveToFolderCommand(
            MoveMultipleFoldersRefactoringAction moveFolders,
            RefactoringUserInteraction<IMoveMultipleFoldersPresenter, MoveMultipleFoldersModel> moveFoldersInteraction,
            MoveMultipleToFolderRefactoringAction moveToFolder,
            RefactoringUserInteraction<IMoveMultipleToFolderPresenter, MoveMultipleToFolderModel> moveToFolderInteraction,
            MoveToFolderRefactoringFailedNotifier failureNotifier, 
            IParserStatusProvider parserStatusProvider, 
            IVbeEvents vbeEvents,
            RubberduckParserState state) 
            : base(moveFolders, moveToFolder, failureNotifier, parserStatusProvider, vbeEvents, state)
        {
            _moveFoldersInteraction = moveFoldersInteraction;
            _moveToFolderInteraction = moveToFolderInteraction;

            AddToCanExecuteEvaluation(SpecialEvaluateCanExecute);
        }

        public override IEnumerable<Type> ApplicableNodeTypes => ApplicableBaseNodes;

        private bool SpecialEvaluateCanExecute(object parameter)
        {
            return parameter is CodeExplorerCustomFolderViewModel
                    || parameter is CodeExplorerComponentViewModel componentViewModel
                        && componentViewModel.Declaration is ModuleDeclaration;
        }

        protected override ICodeExplorerNode NodeFromParameter(object parameter)
        {
            return parameter as ICodeExplorerNode;
        }

        protected override MoveMultipleFoldersModel ModifiedFolderModel(MoveMultipleFoldersModel model, object parameter)
        {
            return _moveFoldersInteraction.UserModifiedModel(model);
        }

        protected override MoveMultipleToFolderModel ModifiedComponentModel(MoveMultipleToFolderModel model, object parameter)
        {
            return _moveToFolderInteraction.UserModifiedModel(model);
        }
    }
}
