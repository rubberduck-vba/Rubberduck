using Rubberduck.Navigation.CodeExplorer;
using Rubberduck.Parsing.VBA;
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

        private readonly IParserStatusProvider _state;
        private readonly ExtractInterfaceRefactoring _refactoring;
        private readonly ExtractInterfaceFailedNotifier _failureNotifier;

        public CodeExplorerExtractInterfaceCommand(
            ExtractInterfaceRefactoring refactoring,
            IParserStatusProvider state,
            ExtractInterfaceFailedNotifier failureNotifier,
            IVbeEvents vbeEvents) 
            : base(vbeEvents)
        {
            _state = state;
            _refactoring = refactoring;
            _failureNotifier = failureNotifier;
            AddToCanExecuteEvaluation(SpecialEvaluateCanExecute);
        }

        public sealed override IEnumerable<Type> ApplicableNodeTypes => ApplicableNodes;

        private bool SpecialEvaluateCanExecute(object parameter)
        {
            return _state.Status == ParserState.Ready &&
                   parameter is CodeExplorerComponentViewModel node &&
                   //node.Declaration.DeclarationType.HasFlag(DeclarationType.ClassModule) &&
                   //node.Children.Any(child => child.Declaration.DeclarationType.HasFlag(DeclarationType.Member));
                   ExtractInterfaceRefactoring.CanExecute((RubberduckParserState)_state, node.QualifiedSelection.Value.QualifiedName);
        }

        protected override void OnExecute(object parameter)
        {
            if (_state.Status != ParserState.Ready ||
                !(parameter is CodeExplorerItemViewModel node) ||
                node.Declaration == null)
            {
                return;
            }

            _refactoring.Refactor(node.Declaration);
        }
    }
}
