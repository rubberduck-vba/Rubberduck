using System;
using System.Collections.Generic;
using Rubberduck.Navigation.CodeExplorer;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.MoveFolder;
using Rubberduck.Refactorings.MoveToFolder;
using Rubberduck.UI.CodeExplorer.Commands.Abstract;
using Rubberduck.VBEditor.Events;
using Rubberduck.UI.Command.Refactorings.Notifiers;

namespace Rubberduck.UI.CodeExplorer.Commands.DragAndDrop
{
    public sealed class CodeExplorerMoveToFolderDragAndDropCommand : CodeExplorerMoveToFolderCommandBase
    {
        public CodeExplorerMoveToFolderDragAndDropCommand(
            MoveMultipleFoldersRefactoringAction moveFolders,
            MoveMultipleToFolderRefactoringAction moveToFolder,
            MoveToFolderRefactoringFailedNotifier failureNotifier, 
            IParserStatusProvider parserStatusProvider, 
            IVbeEvents vbeEvents) 
            : base(moveFolders, moveToFolder, failureNotifier, parserStatusProvider, vbeEvents)
        {
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

        protected override MoveMultipleFoldersModel ModifiedFolderModel(MoveMultipleFoldersModel model, object parameter)
        {
            var (targetFolder, node) = (ValueTuple<string, ICodeExplorerNode>)parameter;
            model.TargetFolder = targetFolder;
            return model;
        }

        protected override MoveMultipleToFolderModel ModifiedComponentModel(MoveMultipleToFolderModel model, object parameter)
        {
            var (targetFolder, node) = (ValueTuple<string, ICodeExplorerNode>)parameter;
            model.TargetFolder = targetFolder;
            return model;
        }
    }
}
