using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.DeleteDeclarations.Abstract;
using System;
using System.Linq;

namespace Rubberduck.Refactorings.DeleteDeclarations
{
    public class DeleteModuleElementsRefactoringAction : DeleteElementsRefactoringActionBase<DeleteModuleElementsModel>
    {
        public DeleteModuleElementsRefactoringAction(IDeclarationFinderProvider declarationFinderProvider, 
            IDeclarationDeletionTargetFactory targetFactory,
            IDeclarationDeletionGroupsGeneratorFactory deletionGroupsGeneratorFactory,
            IRewritingManager rewritingManager)
            : base(declarationFinderProvider, targetFactory, deletionGroupsGeneratorFactory, rewritingManager)
        {}

        public override void Refactor(DeleteModuleElementsModel model, IRewriteSession rewriteSession)
        {
            if (!CanRefactorAllTargets(model))
            {
                throw new InvalidOperationException("Only module-scope declarations can be refactored by this object");
            }

            DeleteDeclarations(model, rewriteSession, base.CreateDeletionTargetsSupportingPartialDeletions);
        }

        protected override bool CanRefactorAllTargets(DeleteModuleElementsModel model)
            => model.Targets.All(t => t.ParentDeclaration is ModuleDeclaration);

        protected override VBAParser.EndOfStatementContext GetPrecedingNonDeletedEOSContextForGroup(IDeclarationDeletionGroup deletionGroup)
        {
            var deleteTarget = deletionGroup.Targets.FirstOrDefault();
            if (!(deleteTarget is IModuleElementDeletionTarget))
            {
                throw new ArgumentException();
            }

            return deletionGroup.PrecedingNonDeletedContext?.GetFollowingEndOfStatementContext()
                ?? deleteTarget.TargetContext.GetAncestor<VBAParser.ModuleContext>()
                    .GetChild<VBAParser.EndOfStatementContext>();
        }
    }
}
