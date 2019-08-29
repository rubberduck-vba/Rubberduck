using System;
using System.Collections.Generic;
using Rubberduck.Navigation.CodeExplorer;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings;
using Rubberduck.Refactorings.Exceptions;
using Rubberduck.Refactorings.Rename;
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
            typeof(CodeExplorerMemberViewModel)
        };

        private readonly IParserStatusProvider _parserStatusProvider;
        private readonly IRefactoring _refactoring;
        private readonly IRefactoringFailureNotifier _failureNotifier;

        public RenameCommand(
            RenameRefactoring refactoring, 
            RenameFailedNotifier renameFailedNotifier, 
            IParserStatusProvider parserStatusProvider, 
            IVbeEvents vbeEvents) 
            : base(vbeEvents)
        {
            _refactoring = refactoring;
            _failureNotifier = renameFailedNotifier;
            _parserStatusProvider = parserStatusProvider;

            AddToCanExecuteEvaluation(SpecialEvaluateCanExecute);
        }

        public override IEnumerable<Type> ApplicableNodeTypes => ApplicableNodes;

        private bool SpecialEvaluateCanExecute(object parameter)
        {
            return _parserStatusProvider.Status == ParserState.Ready;
        }

        protected override void OnExecute(object parameter)
        {
            if (!CanExecute(parameter) ||
                !(parameter is CodeExplorerItemViewModel node) ||
                node.Declaration == null)
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
