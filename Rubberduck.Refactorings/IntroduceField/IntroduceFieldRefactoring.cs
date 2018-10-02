using System;
using System.Collections.Generic;
using System.Linq;
using Rubberduck.Common;
using Rubberduck.Interaction;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.UI;
using Rubberduck.Resources;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.Refactorings.IntroduceField
{
    public class IntroduceFieldRefactoring : IRefactoring
    {
        private readonly IList<Declaration> _declarations;
        private readonly IVBE _vbe;
        private readonly RubberduckParserState _state;
        private readonly IMessageBox _messageBox;

        public IntroduceFieldRefactoring(IVBE vbe, RubberduckParserState state, IMessageBox messageBox)
        {
            _declarations = state.AllUserDeclarations
                .Where(i => i.DeclarationType == DeclarationType.Variable)
                .ToList();

            _vbe = vbe;
            _state = state;
            _messageBox = messageBox;
        }

        public void Refactor()
        {
            var selection = _vbe.GetActiveSelection();

            if (!selection.HasValue)
            {
                _messageBox.NotifyWarn(RubberduckUI.PromoteVariable_InvalidSelection, RubberduckUI.IntroduceField_Caption);
                return;
            }

            Refactor(selection.Value);
        }

        public void Refactor(QualifiedSelection selection)
        {
            var target = _declarations.FindVariable(selection);

            if (target == null)
            {
                _messageBox.NotifyWarn(RubberduckUI.PromoteVariable_InvalidSelection, RubberduckUI.IntroduceParameter_Caption);
                return;
            }

            var rewriter = _state.GetRewriter(target);
            PromoteVariable(rewriter, target);
        }

        public void Refactor(Declaration target)
        {
            if (target.DeclarationType != DeclarationType.Variable)
            {
                _messageBox.NotifyWarn(RubberduckUI.PromoteVariable_InvalidSelection, RubberduckUI.IntroduceParameter_Caption);
                throw new ArgumentException(@"Invalid declaration type", nameof(target));
            }

            var rewriter = _state.GetRewriter(target);
            PromoteVariable(rewriter, target);
        }

        private void PromoteVariable(IModuleRewriter rewriter, Declaration target)
        {
            if (new[] { DeclarationType.ClassModule, DeclarationType.ProceduralModule }.Contains(target.ParentDeclaration.DeclarationType))
            {
                _messageBox.NotifyWarn(RubberduckUI.PromoteVariable_InvalidSelection, RubberduckUI.IntroduceParameter_Caption);
                return;
            }

            var oldSelection = _vbe.GetActiveSelection();

            rewriter.Remove(target);
            AddField(rewriter, target);

            if (oldSelection.HasValue)
            {
                var component = _state.ProjectsProvider.Component(oldSelection.Value.QualifiedName);
                using (var module = component.CodeModule)
                {
                    using (var pane = module.CodePane)
                    {
                        pane.Selection = oldSelection.Value.Selection;
                    }
                }
            }

            rewriter.Rewrite();
        }

        private void AddField(IModuleRewriter rewriter, Declaration target)
        {
            var content = $"{Tokens.Private} {target.IdentifierName} {Tokens.As} {target.AsTypeName}\r\n";
            var members = _state.DeclarationFinder.Members(target.QualifiedName.QualifiedModuleName)
                .Where(item => item.DeclarationType.HasFlag(DeclarationType.Member))
                .OrderByDescending(item => item.Selection);

            var firstMember = members.FirstOrDefault();
            rewriter.InsertBefore(firstMember?.Context.Start.TokenIndex ?? 0, content);
        }
    }
}
