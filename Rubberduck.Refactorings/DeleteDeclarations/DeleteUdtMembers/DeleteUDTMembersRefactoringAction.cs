using Antlr4.Runtime;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using System;
using System.Collections.Generic;
using System.Linq;

namespace Rubberduck.Refactorings.DeleteDeclarations
{
    public class DeleteUDTMembersRefactoringAction : DeleteElementRefactoringActionBase<DeleteUDTMembersModel>
    {
        public DeleteUDTMembersRefactoringAction(IDeclarationFinderProvider declarationFinderProvider, IRewritingManager rewritingManager)
            : base(declarationFinderProvider, rewritingManager)
        {

        }

        public override void Refactor(DeleteUDTMembersModel model, IRewriteSession rewriteSession)
        {
            if (model.Targets.Any(t => t.DeclarationType != DeclarationType.UserDefinedTypeMember))
            {
                throw new InvalidOperationException("Only DeclarationType.UserDefinedTypeMember can be refactored by this class");
            }
            var udtDeleteTargets = new List<IDeclarationDeletionTarget>();

            foreach (var t in model.Targets)
            {
                var udtDeleteTarget = new UdtMemberDeletionTarget(_declarationFinderProvider, t) as IDeclarationDeletionTarget;
                udtDeleteTargets.Add(udtDeleteTarget);
            }

            var udtTargetsByQMN = udtDeleteTargets.GroupBy(t => t.TargetProxy.QualifiedModuleName);

            foreach (var udtGroup in udtTargetsByQMN)
            {
                foreach (var eg in udtGroup)
                {
                    var rewriter = rewriteSession.CheckOutModuleRewriter(udtGroup.Key);
                    DeleteTargetsInModule(udtDeleteTargets, rewriter);
                }
            }
        }

        private void DeleteTargetsInModule(IEnumerable<IDeclarationDeletionTarget> allTargets, IModuleRewriter rewriter)
        {
            var udtMemberDeleteTargets = allTargets
                .Where(dt => dt is IUdtMemberDeletionTarget)
                .ToList();

            if (!udtMemberDeleteTargets.Any())
            {
                return;
            }

            DeleteTargetsInModuleImpl<VBAParser.UdtMemberListContext, VBAParser.UdtMemberContext>(udtMemberDeleteTargets, rewriter);
        }

        protected override List<DeletionGroup> CreateDeletionGroups(List<IDeclarationDeletionTarget> deleteDeclarationTargets, IOrderedEnumerable<ParserRuleContext> udtMembers)
        {
            var nonDeleteIndices = GetNonDeleteContextIndices(udtMembers, deleteDeclarationTargets);

            if (!nonDeleteIndices.Any())
            {
                throw new InvalidOperationException("At least one UDT Member must be retained");
            }

            return AssociatePrecedingNonDeleteContexts(nonDeleteIndices, udtMembers, deleteDeclarationTargets);
        }
    }
}
