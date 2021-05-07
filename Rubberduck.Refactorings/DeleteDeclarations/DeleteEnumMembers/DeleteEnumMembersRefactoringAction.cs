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
    public class DeleteEnumMembersRefactoringAction : DeleteElementRefactoringActionBase<DeleteEnumMembersModel>// CodeOnlyRefactoringActionBase<DeleteEnumMembersModel>
    {
        public DeleteEnumMembersRefactoringAction(IDeclarationFinderProvider declarationFinderProvider, IRewritingManager rewritingManager)
            : base(declarationFinderProvider, rewritingManager)
        {
        }

        public override void Refactor(DeleteEnumMembersModel model)
        {
            throw new NotImplementedException();
        }

        public override void Refactor(DeleteEnumMembersModel model, IRewriteSession rewriteSession)
        {
            if (model.Targets.Any(t => t.DeclarationType != DeclarationType.EnumerationMember))
            {
                throw new InvalidOperationException("Only DeclarationType.EnumerationMember can be refactored by this class");
            }

            var enumDeleteTargets = new List<IDeclarationDeletionTarget>();

            foreach (var t in model.Targets)
            {
                var enumDeleteTarget = new EnumMemberDeletionTarget(_declarationFinderProvider, t) as IDeclarationDeletionTarget;
                enumDeleteTargets.Add(enumDeleteTarget);
            }

            var enumTargetsByQMN = enumDeleteTargets.GroupBy(t => t.TargetProxy.QualifiedModuleName);

            foreach (var enumGroup in enumTargetsByQMN)
            {
                foreach (var eg in enumGroup)
                {
                    var rewriter = rewriteSession.CheckOutModuleRewriter(enumGroup.Key);
                    DeleteTargetsInModule(enumDeleteTargets, rewriter);
                }
            }
        }
        public void DeleteTargetsInModule(IEnumerable<IDeclarationDeletionTarget> allTargets, IModuleRewriter rewriter)
        {
            var enumMemberDeleteTargets = allTargets
                .Where(dt => dt is IEnumMemberDeletionTarget)
                .ToList();

            if (!enumMemberDeleteTargets.Any())
            {
                return;
            }

            DeleteTargetsInModuleImpl<VBAParser.EnumerationStmtContext, VBAParser.EnumerationStmt_ConstantContext>(enumMemberDeleteTargets, rewriter);
        }

        protected override List<DeletionGroup> CreateDeletionGroups(List<IDeclarationDeletionTarget> deleteDeclarationTargets, IOrderedEnumerable<ParserRuleContext> enumMembers)
        {
            var nonDeleteIndices = GetNonDeleteContextIndices(enumMembers, deleteDeclarationTargets);

            if (!nonDeleteIndices.Any())
            {
                throw new InvalidOperationException("At least one Enum Member must be retained");
            }

            return AssociatePrecedingNonDeleteContexts(nonDeleteIndices, enumMembers, deleteDeclarationTargets);
        }
    }
}
