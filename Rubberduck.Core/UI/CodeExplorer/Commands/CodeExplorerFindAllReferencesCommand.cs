using System;
using System.Collections.Generic;
using Rubberduck.AddRemoveReferences;
using Rubberduck.Navigation.CodeExplorer;
using Rubberduck.Parsing.Symbols;
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
        private readonly FindAllReferencesAction _finder;

        public CodeExplorerFindAllReferencesCommand(
            RubberduckParserState state, 
            FindAllReferencesAction finder, 
            IVbeEvents vbeEvents) 
            : base(vbeEvents)
        {
            _state = state;
            _finder = finder;

            AddToCanExecuteEvaluation(EvaluateCanExecute, true);
        }

        private bool EvaluateCanExecute(object parameter)
        {
            switch (parameter)
            {
                case CodeExplorerReferenceViewModel refNode:
                    return refNode.IsDimmed;
                case ICodeExplorerNode node:
                    return !(node is CodeExplorerCustomFolderViewModel)
                        && !(node is CodeExplorerReferenceFolderViewModel);
                case Declaration declaration:
                    return !(declaration is ProjectDeclaration);
                default:
                    return false;
            }
        }

        protected override void OnExecute(object parameter)
        {
            var node = parameter as ICodeExplorerNode;
            var declaration = parameter as Declaration;
            var reference = parameter as CodeExplorerReferenceViewModel;

            if (_state.Status != ParserState.Ready || node == null && declaration == null)
            {
                return;
            }

            if (declaration != null)
            {
                // command could have been invoked from PeekReferences code explorer popup
                _finder.FindAllReferences(declaration);
                return;
            }

            if (reference != null)
            {
                if (!(node.Parent.Declaration is ProjectDeclaration))
                {
                    Logger.Error(
                        $"The specified ICodeExplorerNode ({node.GetType()}) is expected to be a direct child of a node whose declaration is a ProjectDeclaration.");
                    return;
                }

                if(node.Parent?.Declaration is ProjectDeclaration projectDeclaration)
                {
                    if (!(reference.Reference is ReferenceModel model))
                    {
                        Logger.Warn($"Project reference '{reference.Name}' does not have an explorable reference model ({nameof(CodeExplorerReferenceViewModel)}.{nameof(CodeExplorerReferenceViewModel.Reference)} is null.");
                        return;
                    }

                    _finder.FindAllReferences(projectDeclaration, model.ToReferenceInfo());
                    return;
                }
            }

            _finder.FindAllReferences(node.Declaration);
        }

        public override IEnumerable<Type> ApplicableNodeTypes => ApplicableNodes;
    }
}