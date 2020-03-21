using System;
using System.Collections.Generic;
using System.Linq;
using Rubberduck.JunkDrawer.Extensions;
using Rubberduck.Navigation.CodeExplorer;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings;
using Rubberduck.Refactorings.Exceptions;
using Rubberduck.Refactorings.MoveFolder;
using Rubberduck.Refactorings.MoveToFolder;
using Rubberduck.VBEditor.Events;
using Rubberduck.UI.Command.Refactorings.Notifiers;

namespace Rubberduck.UI.CodeExplorer.Commands
{
    public sealed class CodeExplorerMoveToFolderCommand : CodeExplorerCommandBase
    {
        private static readonly Type[] ApplicableNodes =
        {
            typeof(CodeExplorerCustomFolderViewModel),
            typeof(CodeExplorerComponentViewModel)
        };

        private readonly IParserStatusProvider _parserStatusProvider;

        private readonly IRefactoringAction<MoveMultipleFoldersModel> _moveFolders;
        private readonly IRefactoringUserInteraction<MoveMultipleFoldersModel> _moveFoldersInteraction;

        private readonly IRefactoringAction<MoveMultipleToFolderModel> _moveToFolder;
        private readonly IRefactoringUserInteraction<MoveMultipleToFolderModel> _moveToFolderInteraction;

        private readonly IRefactoringFailureNotifier _failureNotifier;

        public CodeExplorerMoveToFolderCommand(
            MoveMultipleFoldersRefactoringAction moveFolders,
            RefactoringUserInteraction<IMoveMultipleFoldersPresenter, MoveMultipleFoldersModel> moveFoldersInteraction,
            MoveMultipleToFolderRefactoringAction moveToFolder,
            RefactoringUserInteraction<IMoveMultipleToFolderPresenter, MoveMultipleToFolderModel> moveToFolderInteraction,
            MoveToFolderRefactoringFailedNotifier failureNotifier, 
            IParserStatusProvider parserStatusProvider, 
            IVbeEvents vbeEvents) 
            : base(vbeEvents)
        {
            _moveFolders = moveFolders;
            _moveFoldersInteraction = moveFoldersInteraction;
            _moveToFolder = moveToFolder;
            _moveToFolderInteraction = moveToFolderInteraction;

            _parserStatusProvider = parserStatusProvider;
            _failureNotifier = failureNotifier;

            AddToCanExecuteEvaluation(SpecialEvaluateCanExecute);
        }

        public override IEnumerable<Type> ApplicableNodeTypes => ApplicableNodes;

        private bool SpecialEvaluateCanExecute(object parameter)
        {
            return _parserStatusProvider.Status == ParserState.Ready
                && (parameter is CodeExplorerCustomFolderViewModel
                    || parameter is CodeExplorerComponentViewModel componentViewModel
                        && componentViewModel.Declaration is ModuleDeclaration);
        }

        protected override void OnExecute(object parameter)
        {
            if (!CanExecute(parameter))
            {
                return;
            }

            if (parameter is CodeExplorerComponentViewModel componentViewModel)
            {
                var model = ComponentModel(componentViewModel);
                var modifiedModel = _moveToFolderInteraction.UserModifiedModel(model);
                ExecuteRefactoringAction(modifiedModel, _moveToFolder, _failureNotifier);
            }

            if (parameter is CodeExplorerCustomFolderViewModel folderViewModel)
            {
                var model = FolderModel(folderViewModel);
                var modifiedModel = _moveFoldersInteraction.UserModifiedModel(model);
                ExecuteRefactoringAction(modifiedModel, _moveFolders, _failureNotifier);
            }
        }

        private MoveMultipleFoldersModel FolderModel(CodeExplorerCustomFolderViewModel folderModel)
        {
            var folder = folderModel.FullPath;
            var containedModules = ContainedModules(folderModel);
            var modulesBySourceFolder = new Dictionary<string, ICollection<ModuleDeclaration>>{{folder, containedModules}};
            var initialTargetFolder = folder.ParentFolder();
            return new MoveMultipleFoldersModel(modulesBySourceFolder, initialTargetFolder);
        }

        private static ICollection<ModuleDeclaration> ContainedModules(ICodeExplorerNode itemModel)
        {
            if (itemModel is CodeExplorerComponentViewModel componentModel)
            {
                var component = componentModel.Declaration;
                return component is ModuleDeclaration moduleDeclaration
                    ? new List<ModuleDeclaration> {moduleDeclaration}
                    : new List<ModuleDeclaration>();
            }

            return itemModel.Children
                .SelectMany(ContainedModules)
                .ToList();
        }

        private MoveMultipleToFolderModel ComponentModel(CodeExplorerComponentViewModel componentViewModel)
        {
            if (!(componentViewModel.Declaration is ModuleDeclaration moduleDeclaration))
            {
                return null;
            }

            var targets = new List<ModuleDeclaration>{moduleDeclaration};
            var targetFolder = moduleDeclaration.CustomFolder;
            return new MoveMultipleToFolderModel(targets, targetFolder);
        } 

        private static void ExecuteRefactoringAction<TModel>(TModel model, IRefactoringAction<TModel> refactoringAction, IRefactoringFailureNotifier failureNotifier)
            where TModel : class, IRefactoringModel
        {
            try
            {
                refactoringAction.Refactor(model);
            }
            catch (RefactoringAbortedException)
            {}
            catch (RefactoringException exception)
            {
                failureNotifier.Notify(exception);
            }
        }
    }
}
