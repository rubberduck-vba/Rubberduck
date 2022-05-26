using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.DeleteDeclarations.Abstract;
using System;
using System.Linq;

namespace Rubberduck.Refactorings.DeleteDeclarations
{
    public class DeleteUDTMembersRefactoringAction : DeleteElementsRefactoringActionBase<DeleteUDTMembersModel>
    {
        public DeleteUDTMembersRefactoringAction(IDeclarationFinderProvider declarationFinderProvider, 
            IDeclarationDeletionTargetFactory targetFactory,
            IDeclarationDeletionGroupsGeneratorFactory deletionGroupsGeneratorFactory,
            IRewritingManager rewritingManager)
            : base(declarationFinderProvider, targetFactory, deletionGroupsGeneratorFactory, rewritingManager)
        {}

        public override void Refactor(DeleteUDTMembersModel model, IRewriteSession rewriteSession)
        {
            if (!CanRefactorAllTargets(model))
            {
                throw new InvalidOperationException("Only DeclarationType.UserDefinedTypeMember can be refactored by this class");
            }

            DeleteDeclarations(model, rewriteSession, (declarations, rewriterSession, targetFactory) => targetFactory.CreateMany(declarations, rewriteSession));
        }

        protected override bool CanRefactorAllTargets(DeleteUDTMembersModel model)
             => model.Targets.Any(t => t.DeclarationType == DeclarationType.UserDefinedTypeMember);
    }
}
