using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.DeleteDeclarations.Abstract;
using System;
using System.Linq;

namespace Rubberduck.Refactorings.DeleteDeclarations
{
    public class DeleteUDTMembersRefactoringAction : DeleteElementRefactoringActionBase<DeleteUDTMembersModel>
    {
        public DeleteUDTMembersRefactoringAction(IDeclarationFinderProvider declarationFinderProvider, 
            IDeclarationDeletionTargetFactory targetFactory,
            IDeclarationDeletionGroupsGeneratorFactory deletionGroupsGeneratorFactory,
            IRewritingManager rewritingManager)
            : base(declarationFinderProvider, targetFactory, deletionGroupsGeneratorFactory, rewritingManager)
        {}

        public override void Refactor(DeleteUDTMembersModel model, IRewriteSession rewriteSession)
        {
            if (model.Targets.Any(t => t.DeclarationType != DeclarationType.UserDefinedTypeMember))
            {
                throw new InvalidOperationException("Only DeclarationType.UserDefinedTypeMember can be refactored by this class");
            }

            DeleteDeclarations(model, rewriteSession);
        }
    }
}
