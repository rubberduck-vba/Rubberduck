using System;
using System.Collections.Generic;
using Rubberduck.Navigation.CodeExplorer;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings;
using Rubberduck.Refactorings.Exceptions;
using Rubberduck.Refactorings.Rename;
using Rubberduck.UI.Command;
using Rubberduck.VBEditor.Events;
using Rubberduck.UI.Command.Refactorings.Notifiers;

namespace Rubberduck.UI.CodeExplorer.Commands
{
    public sealed class RenameCommand : CodeExplorerCommandBase
    {
        private static readonly Type[] ApplicableNodes =
        {
            typeof(CodeExplorerProjectViewModel),
            typeof(CodeExplorerComponentViewModel),
            typeof(CodeExplorerMemberViewModel),
            typeof(CodeExplorerCustomFolderViewModel)
        };

        private readonly IParserStatusProvider _parserStatusProvider;
        private readonly IRefactoring _refactoring;
        private readonly IRefactoringFailureNotifier _failureNotifier;

        private readonly CommandBase _renameFolderCommand;

        public RenameCommand(
            RenameRefactoring refactoring, 
            RenameFailedNotifier renameFailedNotifier, 
            IParserStatusProvider parserStatusProvider, 
            IVbeEvents vbeEvents,
            RenameFolderCommand renameFolderCommand) 
            : base(vbeEvents)
        {
            _refactoring = refactoring;
            _failureNotifier = renameFailedNotifier;
            _parserStatusProvider = parserStatusProvider;
            _renameFolderCommand = renameFolderCommand;

            AddToCanExecuteEvaluation(SpecialEvaluateCanExecute);
        }

        public override IEnumerable<Type> ApplicableNodeTypes => ApplicableNodes;

        private bool SpecialEvaluateCanExecute(object parameter)
        {
            return _parserStatusProvider.Status == ParserState.Ready
                && (!(parameter is CodeExplorerCustomFolderViewModel folderModel)
                    || _renameFolderCommand.CanExecute(folderModel));
        }

        protected override void OnExecute(object parameter)
        {
            if (!CanExecute(parameter) ||
                !(parameter is CodeExplorerItemViewModel node))
            {
                return;
            }

            if (node is CodeExplorerCustomFolderViewModel folderNode)
            {
                _renameFolderCommand.Execute(folderNode);
                return;
            }

            if (node.Declaration == null)
            {
                return;
            }

            try
            {
                _refactoring.Refactor(node.Declaration);
            }
            catch (RefactoringAbortedException)
            { }
            catch (RefactoringException exception)
            {
                _failureNotifier.Notify(exception);
            }
        }
    }
}
