using System;
using System.Collections.Generic;
using System.Linq;
using Rubberduck.Interaction;
using Rubberduck.JunkDrawer.Extensions;
using Rubberduck.Navigation.CodeExplorer;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.Exceptions;
using Rubberduck.Refactorings.MoveFolder;
using Rubberduck.Refactorings.MoveToFolder;
using Rubberduck.Resources;
using Rubberduck.UI.CodeExplorer.Commands.Abstract;
using Rubberduck.VBEditor.Events;
using Rubberduck.UI.Command.Refactorings.Notifiers;

namespace Rubberduck.UI.CodeExplorer.Commands.DragAndDrop
{
    public sealed class CodeExplorerMoveToFolderDragAndDropCommand : CodeExplorerMoveToFolderCommandBase
    {
        private readonly IDeclarationFinderProvider _declarationFinderProvider;
        private readonly IMessageBox _messageBox;

        public CodeExplorerMoveToFolderDragAndDropCommand(
            MoveMultipleFoldersRefactoringAction moveFolders,
            MoveMultipleToFolderRefactoringAction moveToFolder,
            MoveToFolderRefactoringFailedNotifier failureNotifier, 
            IParserStatusProvider parserStatusProvider, 
            IVbeEvents vbeEvents,
            IMessageBox messageBox,
            IDeclarationFinderProvider declarationFinderProvider,
            RubberduckParserState state) 
            : base(moveFolders, moveToFolder, failureNotifier, parserStatusProvider, vbeEvents, state)
        {
            _declarationFinderProvider = declarationFinderProvider;
            _messageBox = messageBox;

            AddToCanExecuteEvaluation(SpecialEvaluateCanExecute);
        }

        //We need to use the version with the interface since the type parameters always have to match exactly in the check in the base class.
        public override IEnumerable<Type> ApplicableNodeTypes => new []{typeof(ValueTuple<string, ICodeExplorerNode>)};

        private bool SpecialEvaluateCanExecute(object parameter)
        {
            var (targetFolder, node) = (ValueTuple<string, ICodeExplorerNode>)parameter;
            return !string.IsNullOrEmpty(targetFolder)
                   && (node is CodeExplorerCustomFolderViewModel folderViewModel
                       && folderViewModel.FullPath != targetFolder
                    || node is CodeExplorerComponentViewModel componentViewModel
                       && componentViewModel.Declaration is ModuleDeclaration);
        }

        protected override ICodeExplorerNode NodeFromParameter(object parameter)
        {
            var (targetFolder, node) = (ValueTuple<string, ICodeExplorerNode>)parameter;
            return node;
        }

        protected override MoveMultipleToFolderModel ModifiedComponentModel(MoveMultipleToFolderModel model, object parameter)
        {
            var (targetFolder, node) = (ValueTuple<string, ICodeExplorerNode>)parameter;
            model.TargetFolder = targetFolder;
            return model;
        }

        protected override MoveMultipleFoldersModel ModifiedFolderModel(MoveMultipleFoldersModel model, object parameter)
        {
            var (targetFolder, node) = (ValueTuple<string, ICodeExplorerNode>)parameter;
            if (OkToMoveFolders(model, targetFolder))
            {
                model.TargetFolder = targetFolder;
            }
            else
            {
                throw new RefactoringAbortedException();
            }

            return model;
        }

        private bool OkToMoveFolders(MoveMultipleFoldersModel model, string targetFolder)
        {
            var foldersMergedWithTargetFolders = FoldersMergedWithTargetFolders(model, targetFolder);
            return !foldersMergedWithTargetFolders.Any()
                   || UserConfirmsToProceedWithFolderMerge(targetFolder, foldersMergedWithTargetFolders);
        }

        private List<string> FoldersMergedWithTargetFolders(MoveMultipleFoldersModel model, string targetFolder)
        {
            var movingFolders = model.ModulesBySourceFolder
                .Select(kvp => kvp.Key)
                .Where(folder => !folder.ParentFolder().Equals(targetFolder))
                .Select(folder => folder.SubFolderName());

            var targetFolderSubfolders = _declarationFinderProvider.DeclarationFinder
                .UserDeclarations(DeclarationType.Module)
                .OfType<ModuleDeclaration>()
                .Select(module => module.CustomFolder)
                .Where(folder => folder.IsSubFolderOf(targetFolder))
                .Select(folder => folder.SubFolderPathRelativeTo(targetFolder).RootFolder())
                .ToHashSet();

            return movingFolders
                .Where(folder => targetFolderSubfolders.Contains(folder))
                .ToList();
        }

        private bool UserConfirmsToProceedWithFolderMerge(string targetFolder, List<string> mergedTargetFolders)
        {
            var message = FolderMergeUserConfirmationMessage(targetFolder, mergedTargetFolders);
            return _messageBox?.ConfirmYesNo(message, RubberduckUI.MoveFoldersDialog_Caption) ?? false;
        }

        private string FolderMergeUserConfirmationMessage(string targetFolder, List<string> mergedTargetFolders)
        {
            if (mergedTargetFolders.Count == 1)
            {
                return string.Format(
                    RubberduckUI.MoveFolders_SameNameSubfolder,
                    targetFolder,
                    mergedTargetFolders.Single());
            }

            var subfolders = $"'{string.Join("', '", mergedTargetFolders)}'";
            return string.Format(
                RubberduckUI.MoveFolders_SameNameSubfolders,
                targetFolder,
                subfolders);
        }
    }
}
