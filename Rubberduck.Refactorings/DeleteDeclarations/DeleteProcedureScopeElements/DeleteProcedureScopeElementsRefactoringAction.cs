using Antlr4.Runtime;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using System;
using System.Collections.Generic;
using System.Linq;

namespace Rubberduck.Refactorings.DeleteDeclarations
{
    public class DeleteProcedureScopeElementsRefactoringAction : DeleteVariableOrConstantRefactoringActionBase<DeleteProcedureScopeElementsModel>
    {
        public DeleteProcedureScopeElementsRefactoringAction(IDeclarationFinderProvider declarationFinderProvider, IDeclarationDeletionTargetFactory targetFactory, IDeleteDeclarationEndOfStatementContentModifierFactory eosModifierFactory, IRewritingManager rewritingManager)
           : base(declarationFinderProvider, targetFactory, eosModifierFactory, rewritingManager)
        {}
        public override void Refactor(DeleteProcedureScopeElementsModel model, IRewriteSession rewriteSession)
        {
            DeleteDeclarations(model, rewriteSession);
        }

        protected override List<IDeclarationDeletionTarget> HandleLabelAndVarOrConstInSameBlock(List<IDeclarationDeletionTarget> blockDeleteTargets)
        {
            var blockDeleteTargetsByContext = blockDeleteTargets.ToLookup(k => k.TargetContext);
            if (blockDeleteTargetsByContext.Count < blockDeleteTargets.Count)
            {
                var newList = new List<IDeclarationDeletionTarget>();
                foreach (var deletionTargets in blockDeleteTargetsByContext)
                {
                    var varOrConst = deletionTargets.OfType<IProcedureLocalDeletionTarget>().Single();
                    varOrConst.DeleteAssociatedLabel = deletionTargets.Any(e => e is ILineLabelDeletionTarget);
                    newList.Add(varOrConst);
                    return newList;
                }
            }
            return blockDeleteTargets;
        }

        protected override IOrderedEnumerable<ParserRuleContext> GetAllContextElements(Declaration declaration)
            => GetAllTargetContextElements<VBAParser.BlockContext, VBAParser.BlockStmtContext>(declaration);

        protected override void RefactorGuardClause(IDeleteDeclarationsModel model)
        {
            if (model.Targets.Any(t => t.ParentDeclaration is ModuleDeclaration))
            {
                throw new InvalidOperationException("Only declarations within procedures can be refactored by this object");
            }
        }
    }
}
