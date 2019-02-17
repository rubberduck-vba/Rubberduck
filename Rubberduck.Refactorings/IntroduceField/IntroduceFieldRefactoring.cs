using System;
using System.Linq;
using Rubberduck.Common;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.Exceptions;
using Rubberduck.Refactorings.Exceptions.IntroduceField;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.Utility;

namespace Rubberduck.Refactorings.IntroduceField
{
    public class IntroduceFieldRefactoring : RefactoringBase
    {
        private readonly IDeclarationFinderProvider _declarationFinderProvider;

        public IntroduceFieldRefactoring(IDeclarationFinderProvider declarationFinderProvider, IRewritingManager rewritingManager, ISelectionService selectionService)
        :base(rewritingManager, selectionService)
        {
            _declarationFinderProvider = declarationFinderProvider;
        }

        public override void Refactor(QualifiedSelection target)
        {
            var targetDeclaration = FindTarget(target);

            if (targetDeclaration == null)
            {
                throw new NoDeclarationForSelectionException(target);
            }

            Refactor(targetDeclaration);
        }

        private Declaration FindTarget(QualifiedSelection targetSelection)
        {
            return _declarationFinderProvider.DeclarationFinder
                .UserDeclarations(DeclarationType.Variable)
                .FindVariable(targetSelection);
        }

        public override void Refactor(Declaration target)
        {
            if (target == null)
            {
                throw new TargetDeclarationIsNullException(target);
            }

            if (target.DeclarationType != DeclarationType.Variable)
            {
                throw new InvalidDeclarationTypeException(target);
            }

            if (new[] { DeclarationType.ClassModule, DeclarationType.ProceduralModule }.Contains(target.ParentDeclaration.DeclarationType))
            {
                throw new TargetIsAlreadyAFieldException(target);
            }

            PromoteVariable(target);
        }

        private void PromoteVariable(Declaration target)
        {
            var rewriteSession = RewritingManager.CheckOutCodePaneSession();
            var rewriter = rewriteSession.CheckOutModuleRewriter(target.QualifiedModuleName);

            rewriter.Remove(target);
            AddField(rewriter, target);

            rewriteSession.TryRewrite();
        }

        private void AddField(IModuleRewriter rewriter, Declaration target)
        {
            var content = $"{Tokens.Private} {target.IdentifierName} {Tokens.As} {target.AsTypeName}{Environment.NewLine}";
            var members = _declarationFinderProvider.DeclarationFinder.Members(target.QualifiedName.QualifiedModuleName)
                .Where(item => item.DeclarationType.HasFlag(DeclarationType.Member))
                .OrderBy(item => item.Selection);

            var firstMember = members.FirstOrDefault();
            rewriter.InsertBefore(firstMember?.Context.Start.TokenIndex ?? 0, content);
        }
    }
}
