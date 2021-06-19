using Antlr4.Runtime;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using System;
using System.Linq;

namespace Rubberduck.Refactorings.DeleteDeclarations
{
    public class DeleteUDTMembersRefactoringAction : DeleteElementRefactoringActionBase<DeleteUDTMembersModel>
    {
        public DeleteUDTMembersRefactoringAction(IDeclarationFinderProvider declarationFinderProvider, IDeclarationDeletionTargetFactory targetFactory, IDeleteDeclarationEndOfStatementContentModifierFactory eosModifierFactory, IRewritingManager rewritingManager)
            : base(declarationFinderProvider, targetFactory, eosModifierFactory, rewritingManager)
        {}

        public override void Refactor(DeleteUDTMembersModel model, IRewriteSession rewriteSession)
        {
            model.RemoveAllExceptionMessage = "At least one UDT Member must be retained";

            DeleteDeclarations(model, rewriteSession);
        }

        protected override void RefactorGuardClause(IDeleteDeclarationsModel model)
        {
            if (model.Targets.Any(t => t.DeclarationType != DeclarationType.UserDefinedTypeMember))
            {
                throw new InvalidOperationException("Only DeclarationType.UserDefinedTypeMember can be refactored by this class");
            }
        }

        protected override IOrderedEnumerable<ParserRuleContext> GetAllContextElements(Declaration declaration)
            => GetAllTargetContextElements<VBAParser.UdtMemberListContext, VBAParser.UdtMemberContext>(declaration);
    }
}
