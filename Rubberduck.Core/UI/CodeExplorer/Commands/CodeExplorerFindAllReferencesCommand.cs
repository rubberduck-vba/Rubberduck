using System;
using System.Collections.Generic;
using Rubberduck.AddRemoveReferences;
using Rubberduck.Navigation.CodeExplorer;
using Rubberduck.Parsing.VBA;
using Rubberduck.UI.Controls;

namespace Rubberduck.UI.CodeExplorer.Commands
{
    public class CodeExplorerFindAllReferencesCommand : CodeExplorerCommandBase
    {
        private static readonly Type[] ApplicableNodes =
        {
            typeof(CodeExplorerReferenceViewModel),
            typeof(CodeExplorerProjectViewModel),
            typeof(CodeExplorerComponentViewModel),
            typeof(CodeExplorerMemberViewModel)
        };

        private readonly RubberduckParserState _state;
        private readonly FindAllReferencesService _finder;

        public CodeExplorerFindAllReferencesCommand(RubberduckParserState state, FindAllReferencesService finder)
        {
            _state = state;
            _finder = finder;
        }

        protected override void OnExecute(object parameter)
        {
            if (_state.Status != ParserState.Ready ||
                !(parameter is ICodeExplorerNode node) ||
                node.Declaration == null)
            {
                return;
            }

            if (parameter is CodeExplorerReferenceViewModel reference)
            {
                if (!(reference.Reference is ReferenceModel model))
                {
                    return;
                }
                _finder.FindAllReferences(node.Parent.Declaration, model.ToReferenceInfo());
                return;
            }

            _finder.FindAllReferences(node.Declaration);
        }

        public override IEnumerable<Type> ApplicableNodeTypes => ApplicableNodes;

        protected override bool EvaluateCanExecute(object parameter)
        {
            return base.EvaluateCanExecute(parameter) && 
                   ((ICodeExplorerNode)parameter).Declaration != null &&
                   (!(parameter is CodeExplorerReferenceViewModel reference) || !reference.IsDimmed) &&
                   _state.Status == ParserState.Ready;
        }
    }
}
