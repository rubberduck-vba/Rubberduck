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
using Rubberduck.VBEditor.Utility;

namespace Rubberduck.Refactorings.IntroduceField
{
    public class IntroduceFieldRefactoring : RefactoringBase
    {
        private readonly IDeclarationFinderProvider _declarationFinderProvider;
        private readonly IMessageBox _messageBox;

        public IntroduceFieldRefactoring(IDeclarationFinderProvider declarationFinderProvider, IMessageBox messageBox, IRewritingManager rewritingManager, ISelectionService selectionService)
        :base(rewritingManager, selectionService)
        {
            _declarationFinderProvider = declarationFinderProvider;
            _messageBox = messageBox;
        }

        public override void Refactor()
        {
            var activeSelection = SelectionService.ActiveSelection();

            if (!activeSelection.HasValue)
            {
                _messageBox.NotifyWarn(RubberduckUI.PromoteVariable_InvalidSelection, RubberduckUI.IntroduceField_Caption);
                return;
            }

            Refactor(activeSelection.Value);
        }

        public override void Refactor(QualifiedSelection selection)
        {
            var target = _declarationFinderProvider.DeclarationFinder
                .UserDeclarations(DeclarationType.Variable)
                .FindVariable(selection);

            if (target == null)
            {
                _messageBox.NotifyWarn(RubberduckUI.PromoteVariable_InvalidSelection, RubberduckUI.IntroduceParameter_Caption);
                return;
            }

            PromoteVariable(target);
        }

        public override void Refactor(Declaration target)
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

            var rewriteSession = RewritingManager.CheckOutCodePaneSession();
            var rewriter = rewriteSession.CheckOutModuleRewriter(target.QualifiedModuleName);

            rewriter.Remove(target);
            AddField(rewriter, target);

            rewriteSession.TryRewrite();
        }

        private void AddField(IModuleRewriter rewriter, Declaration target)
        {
            var content = $"{Tokens.Private} {target.IdentifierName} {Tokens.As} {target.AsTypeName}\r\n";
            var members = _declarationFinderProvider.DeclarationFinder.Members(target.QualifiedName.QualifiedModuleName)
                .Where(item => item.DeclarationType.HasFlag(DeclarationType.Member))
                .OrderByDescending(item => item.Selection);

            var firstMember = members.FirstOrDefault();
            rewriter.InsertBefore(firstMember?.Context.Start.TokenIndex ?? 0, content);
        }
    }
}
