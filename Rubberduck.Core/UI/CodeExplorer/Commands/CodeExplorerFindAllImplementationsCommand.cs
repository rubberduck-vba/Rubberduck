using System;
using System.Collections.Generic;
using Rubberduck.Navigation.CodeExplorer;
using Rubberduck.Parsing.VBA;
using Rubberduck.UI.Controls;

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

        public CodeExplorerFindAllImplementationsCommand(RubberduckParserState state, FindAllImplementationsService finder)
        {
            _state = state;
            _finder = finder;
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

        public override IEnumerable<Type> ApplicableNodeTypes => ApplicableNodes;

        protected override bool EvaluateCanExecute(object parameter)
        {
            return base.EvaluateCanExecute(parameter) &&
                   _state.Status == ParserState.Ready &&
                   parameter is CodeExplorerItemViewModel node &&
                   _finder.CanFind(node.Declaration);
        }
    }
}
