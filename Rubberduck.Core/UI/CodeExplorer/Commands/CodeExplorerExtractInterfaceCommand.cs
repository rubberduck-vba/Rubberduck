using Rubberduck.Navigation.CodeExplorer;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.Exceptions;
using Rubberduck.Refactorings.ExtractInterface;
using Rubberduck.UI.Command.Refactorings.Notifiers;
using Rubberduck.VBEditor.Events;
using System;
using System.Collections.Generic;

namespace Rubberduck.UI.CodeExplorer.Commands
{
    public class CodeExplorerExtractInterfaceCommand : CodeExplorerCommandBase
    {
        private static readonly Type[] ApplicableNodes =
        {
            typeof(CodeExplorerComponentViewModel)
        };

        private readonly RubberduckParserState _state;
        private readonly ExtractInterfaceRefactoring _refactoring;
        private readonly ExtractInterfaceFailedNotifier _failureNotifier;

        public CodeExplorerExtractInterfaceCommand(
            ExtractInterfaceRefactoring refactoring,
            RubberduckParserState state,
            ExtractInterfaceFailedNotifier failureNotifier,
            IVbeEvents vbeEvents) 
            : base(vbeEvents)
        {
            _state = state;
            _refactoring = refactoring;
            _failureNotifier = failureNotifier;
            AddToCanExecuteEvaluation(EvaluateCanExecute);
            AddToOnExecuteEvaluation(EvaluateWillExecute);
        }

        public sealed override IEnumerable<Type> ApplicableNodeTypes => ApplicableNodes;

        private bool EvaluateCanExecute(object parameter)
        {
            return _state.Status == ParserState.Ready 
                   && parameter is CodeExplorerComponentViewModel node 
                   && node.QualifiedSelection.HasValue
                   && _refactoring.CanExecute(_state, node.QualifiedSelection.Value.QualifiedName);
        }

        private bool EvaluateWillExecute(object parameter)
        {
            return _state.Status == ParserState.Ready 
                   && parameter is CodeExplorerItemViewModel node 
                   && node.Declaration != null;
        }

        protected override void OnExecute(object parameter)
        {
            try
            {
                _refactoring.Refactor(((CodeExplorerItemViewModel)parameter).Declaration);
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
