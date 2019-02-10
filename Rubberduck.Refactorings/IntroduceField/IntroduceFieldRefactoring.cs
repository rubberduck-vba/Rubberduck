using System;
using System.Collections.Generic;
using System.Linq;
using Rubberduck.Common;
using Rubberduck.Interaction;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Resources;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using Rubberduck.VBEditor.Utility;

namespace Rubberduck.Refactorings.IntroduceField
{
    public class IntroduceFieldRefactoring : IRefactoring
    {
        private readonly IList<Declaration> _declarations;
        private readonly ISelectionService _selectionService;
        private readonly RubberduckParserState _state;
        private readonly IRewritingManager _rewritingManager;
        private readonly IMessageBox _messageBox;

        public IntroduceFieldRefactoring(RubberduckParserState state, IMessageBox messageBox, IRewritingManager rewritingManager, ISelectionService selectionService)
        {
            _declarations = state.AllUserDeclarations
                .Where(i => i.DeclarationType == DeclarationType.Variable)
                .ToList();

            _state = state;
            _rewritingManager = rewritingManager;
            _messageBox = messageBox;
            _selectionService = selectionService;
        }

        public void Refactor()
        {
            var activeSelection = _selectionService.ActiveSelection();

            if (!activeSelection.HasValue)
            {
                _messageBox.NotifyWarn(RubberduckUI.PromoteVariable_InvalidSelection, RubberduckUI.IntroduceField_Caption);
                return;
            }

            Refactor(activeSelection.Value);
        }

        public void Refactor(QualifiedSelection selection)
        {
            var target = _declarations.FindVariable(selection);

            if (target == null)
            {
                _messageBox.NotifyWarn(RubberduckUI.PromoteVariable_InvalidSelection, RubberduckUI.IntroduceParameter_Caption);
                return;
            }

            PromoteVariable(target);
        }

        public void Refactor(Declaration target)
        {
            if (target.DeclarationType != DeclarationType.Variable)
            {
                _messageBox.NotifyWarn(RubberduckUI.PromoteVariable_InvalidSelection, RubberduckUI.IntroduceParameter_Caption);
                throw new ArgumentException(@"Invalid declaration type", nameof(target));
            }

            PromoteVariable(target);
        }

        private void PromoteVariable(Declaration target)
        {
            if (new[] { DeclarationType.ClassModule, DeclarationType.ProceduralModule }.Contains(target.ParentDeclaration.DeclarationType))
            {
                _messageBox.NotifyWarn(RubberduckUI.PromoteVariable_InvalidSelection, RubberduckUI.IntroduceParameter_Caption);
                return;
            }

            var rewriteSession = _rewritingManager.CheckOutCodePaneSession();
            var rewriter = rewriteSession.CheckOutModuleRewriter(target.QualifiedModuleName);

            rewriter.Remove(target);
            AddField(rewriter, target);

            rewriteSession.TryRewrite();
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
