using System;
using System.Linq;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.Refactorings.IntroduceField
{
    public class IntroduceFieldRefactoringAction : RefactoringActionBase<IntroduceFieldModel>
    {
        private readonly IDeclarationFinderProvider _declarationFinderProvider;

        public IntroduceFieldRefactoringAction(IDeclarationFinderProvider declarationFinderProvider,
            IRewritingManager rewritingManager)
            : base(rewritingManager)
        {
            _declarationFinderProvider = declarationFinderProvider;
        }

        protected override void Refactor(IntroduceFieldModel model, IRewriteSession rewriteSession)
        {
            var target = model.Target;
            var rewriter = rewriteSession.CheckOutModuleRewriter(target.QualifiedModuleName);
            rewriter.Remove(target);
            AddField(rewriter, target);
        }

        private void AddField(IModuleRewriter rewriter, Declaration target)
        {
            var content = $"{Tokens.Private} {target.IdentifierName} {Tokens.As} {target.AsTypeName}{Environment.NewLine}";
            var members = _declarationFinderProvider.DeclarationFinder
                .Members(target.QualifiedName.QualifiedModuleName, DeclarationType.Member)
                .OrderBy(item => item.Selection);

            var firstMember = members.FirstOrDefault();
            rewriter.InsertBefore(firstMember?.Context.Start.TokenIndex ?? 0, content);
        }
    }
}