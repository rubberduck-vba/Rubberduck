using System;
using System.Collections.Generic;
using Rubberduck.AddRemoveReferences;
using Rubberduck.Navigation.CodeExplorer;
using Rubberduck.Parsing.VBA;
using Rubberduck.UI.Controls;
using Rubberduck.VBEditor.Events;

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

        public CodeExplorerFindAllReferencesCommand(
            RubberduckParserState state, 
            FindAllReferencesService finder, 
            IVbeEvents vbeEvents) 
            : base(vbeEvents)
        {
            _state = state;
            _finder = finder;

            AddToCanExecuteEvaluation(SpecialEvaluateCanExecute);
        }

        private bool SpecialEvaluateCanExecute(object parameter)
        {
            return _state.Status == ParserState.Ready
                       && ((ICodeExplorerNode)parameter).Declaration != null 
                       && !(parameter is CodeExplorerReferenceViewModel reference 
                            && reference.IsDimmed);
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
    }
}
