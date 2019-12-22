using System;
using System.Collections.Generic;
using Rubberduck.Navigation.CodeExplorer;
using Rubberduck.Parsing.VBA;
using Rubberduck.UI.Controls;
using Rubberduck.VBEditor.Events;

namespace Rubberduck.UI.CodeExplorer.Commands
{
    public class CodeExplorerFindAllImplementationsCommand : CodeExplorerCommandBase
    {
        private static readonly Type[] ApplicableNodes =
        {
            typeof(CodeExplorerComponentViewModel),
            typeof(CodeExplorerMemberViewModel)
        };

        private readonly RubberduckParserState _state;
        private readonly FindAllImplementationsService _finder;

        public CodeExplorerFindAllImplementationsCommand(
            RubberduckParserState state, 
            FindAllImplementationsService finder, 
            IVbeEvents vbeEvents) 
            : base(vbeEvents)
        {
            _state = state;
            _finder = finder;

            AddToCanExecuteEvaluation(SpecialEvaluateCanExecute);
        }

        public sealed override IEnumerable<Type> ApplicableNodeTypes => ApplicableNodes;

        private bool SpecialEvaluateCanExecute(object parameter)
        {
            return _state.Status == ParserState.Ready &&
                   parameter is CodeExplorerItemViewModel node &&
                   _finder.CanFind(node.Declaration);
        }

        protected override void OnExecute(object parameter)
        {
            if (_state.Status != ParserState.Ready ||
                !(parameter is CodeExplorerItemViewModel node) ||
                node.Declaration == null)
            {
                return;
            }

            _finder.FindAllImplementations(node.Declaration);
        }
    }
}
