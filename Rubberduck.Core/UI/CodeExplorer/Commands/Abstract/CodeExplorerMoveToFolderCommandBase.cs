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

namespace Rubberduck.UI.CodeExplorer.Commands.Abstract
{
    public abstract class CodeExplorerMoveToFolderCommandBase : CodeExplorerCommandBase
    {
        protected static readonly Type[] ApplicableBaseNodes =
        {
            typeof(CodeExplorerCustomFolderViewModel),
            typeof(CodeExplorerComponentViewModel)
        };

        private readonly IParserStatusProvider _parserStatusProvider;
        private readonly RubberduckParserState _state;

        private readonly IRefactoringAction<MoveMultipleFoldersModel> _moveFolders;
        private readonly IRefactoringAction<MoveMultipleToFolderModel> _moveToFolder;

        private readonly IRefactoringFailureNotifier _failureNotifier;

        protected CodeExplorerMoveToFolderCommandBase(
            MoveMultipleFoldersRefactoringAction moveFolders,
            MoveMultipleToFolderRefactoringAction moveToFolder,
            MoveToFolderRefactoringFailedNotifier failureNotifier, 
            IParserStatusProvider parserStatusProvider, 
            IVbeEvents vbeEvents,
            RubberduckParserState state) 
            : base(vbeEvents)
        {
            _moveFolders = moveFolders;
            _moveToFolder = moveToFolder;

            _parserStatusProvider = parserStatusProvider;
            _state = state;
            _failureNotifier = failureNotifier;

            AddToCanExecuteEvaluation(SpecialEvaluateCanExecute);
        }

        private bool SpecialEvaluateCanExecute(object parameter)
        {
            return _parserStatusProvider.Status == ParserState.Ready;
        }

        protected abstract ICodeExplorerNode NodeFromParameter(object parameter);
        protected abstract MoveMultipleFoldersModel ModifiedFolderModel(MoveMultipleFoldersModel model, object parameter);
        protected abstract MoveMultipleToFolderModel ModifiedComponentModel(MoveMultipleToFolderModel model, object parameter);

        protected override void OnExecute(object parameter)
        {
            if (!CanExecute(parameter))
            {
                return;
            }

            var node = NodeFromParameter(parameter);

            if (node is CodeExplorerComponentViewModel componentViewModel)
            {
                var model = ComponentModel(componentViewModel);
                ExecuteRefactoringAction(model, parameter, ValidateInitialComponentModel, ModifiedComponentModel, _moveToFolder, _failureNotifier);
            }

            if (node is CodeExplorerCustomFolderViewModel folderViewModel)
            {
                var model = FolderModel(folderViewModel);
                ExecuteRefactoringAction(model, parameter, ValidateInitialFolderModel, ModifiedFolderModel, _moveFolders, _failureNotifier);
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

        private void ValidateInitialFolderModel(MoveMultipleFoldersModel model)
        {
            var firstStaleAffectedModules = model.ModulesBySourceFolder.Values
                .SelectMany(modules => modules)
                .FirstOrDefault(module => _state.IsNewOrModified(module.QualifiedModuleName));
            if (firstStaleAffectedModules != null)
            {
                throw new AffectedModuleIsStaleException(firstStaleAffectedModules.QualifiedModuleName);
            }
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

        private void ValidateInitialComponentModel(MoveMultipleToFolderModel model)
        {
            var firstStaleAffectedModules = model.Targets
                .FirstOrDefault(module => _state.IsNewOrModified(module.QualifiedModuleName));
            if (firstStaleAffectedModules != null)
            {
                throw new AffectedModuleIsStaleException(firstStaleAffectedModules.QualifiedModuleName);
            }
        }

        private static void ExecuteRefactoringAction<TModel>(
            TModel model, 
            object parameter,
            Action<TModel> initialModelValidation,
            Func<TModel,object,TModel> modelModification, 
            IRefactoringAction<TModel> refactoringAction, 
            IRefactoringFailureNotifier failureNotifier)
            where TModel : class, IRefactoringModel
        {
            try
            {
                initialModelValidation(model);
                var modifiedModel = modelModification(model, parameter);
                refactoringAction.Refactor(modifiedModel);
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
