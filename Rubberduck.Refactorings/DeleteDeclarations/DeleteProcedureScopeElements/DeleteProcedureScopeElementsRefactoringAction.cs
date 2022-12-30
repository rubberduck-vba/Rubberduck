using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.DeleteDeclarations.Abstract;
using System;
using System.Collections.Generic;
using System.Linq;

namespace Rubberduck.Refactorings.DeleteDeclarations
{
    public class DeleteProcedureScopeElementsRefactoringAction : DeleteElementsRefactoringActionBase<DeleteProcedureScopeElementsModel>
    {
        public DeleteProcedureScopeElementsRefactoringAction(IDeclarationFinderProvider declarationFinderProvider, 
            IDeclarationDeletionTargetFactory targetFactory,
            IDeclarationDeletionGroupsGeneratorFactory deletionGroupsGeneratorFactory,
            IRewritingManager rewritingManager)
           : base(declarationFinderProvider, targetFactory, deletionGroupsGeneratorFactory, rewritingManager)
        {
            DeleteTarget = DeleteLocalTarget;
        }

        public override void Refactor(DeleteProcedureScopeElementsModel model, IRewriteSession rewriteSession)
        {
            if (!CanRefactorAllTargets(model))
            {
                throw new InvalidOperationException("Only declarations within procedures other than parameters can be refactored by this object");
            }

            DeleteDeclarations(model, rewriteSession, CreateDeletionTargetsLocalScope);
        }

        protected override bool CanRefactorAllTargets(DeleteProcedureScopeElementsModel model)
            => model.Targets.All(t => t.ParentDeclaration.DeclarationType.HasFlag(DeclarationType.Member) && !t.DeclarationType.HasFlag(DeclarationType.Parameter));

        //Creating local targets are the same as creating a module level deletion targets except that Labels 
        //need to be accounted for
        private IEnumerable<IDeclarationDeletionTarget> CreateDeletionTargetsLocalScope(IEnumerable<Declaration> declarations, IRewriteSession rewriteSession, IDeclarationDeletionTargetFactory targetFactory)
        {
            var allLocalScopeTargets = base.CreateDeletionTargetsSupportingPartialDeletions(declarations, rewriteSession, targetFactory)
                 .Cast<ILocalScopeDeletionTarget>();

            var labelTargets = allLocalScopeTargets.Where(t => t.IsLabel(out _)).Cast<ILabelDeletionTarget>().ToList();
            if (!labelTargets.Any())
            {
                return allLocalScopeTargets;
            }

            var variableOrConstantTargets = allLocalScopeTargets.Except(labelTargets.Cast<ILocalScopeDeletionTarget>()).ToList();

            var labelTargetsDeletedByAssociation = new List<ILabelDeletionTarget>();
            foreach (var labelTarget in labelTargets)
            {
                if (!labelTarget.HasSameLogicalLineListContext(out var relatedVarOrConstListContext))
                {
                    continue;
                }

                //When a Label and Variable/Const Declaration are on the same logical line AND both are to be deleted,
                //associate the Label declaration with the Declaration ListContext.  Deleting the Declaration will 
                //result in deleting the Label as well because the entire BlockStmtContext will be deleted.
                var relatedTarget = variableOrConstantTargets.SingleOrDefault(t => t.ListContext == relatedVarOrConstListContext);
                if (relatedTarget != null)
                {
                    relatedTarget.SetupToDeleteAssociatedLabel(labelTarget);

                    labelTargetsDeletedByAssociation.Add(labelTarget);
                }
            }

            labelTargets.RemoveAll(t => labelTargetsDeletedByAssociation.Contains(t));

            //A deleted Label that has content in the same VBAParser.BlockStmtContext replaces the label expression with equivalent whitespace
            foreach (var labelTarget in labelTargets)
            {
                labelTarget.ReplaceLabelWithWhitespace = labelTarget.HasSameLogicalLineListContext(out var varOrConstListContext)
                    ? !variableOrConstantTargets.Any(t => t.ListContext == varOrConstListContext)
                    : labelTarget.HasFollowingMainBlockStatementContext(out _);
            }

            return variableOrConstantTargets.Concat(labelTargets.Cast<ILocalScopeDeletionTarget>()).ToList();
        }

        private static void DeleteLocalTarget(IDeclarationDeletionTarget deleteTarget, IModuleRewriter rewriter)
        {
            if (deleteTarget is ILabelDeletionTarget lblTarget && lblTarget.ReplaceLabelWithWhitespace)
            {
                var labelContent = deleteTarget.DeleteContext.GetText();
                var spaces = string.Concat(Enumerable.Repeat(' ', labelContent.Length));
                rewriter.Replace(deleteTarget.DeleteContext, spaces);
                return;
            }

            rewriter.Remove(deleteTarget.DeleteContext);
        }

        protected override VBAParser.EndOfStatementContext GetPrecedingNonDeletedEOSContextForGroup(IDeclarationDeletionGroup deletionGroup)
        {
            var firstTarget = deletionGroup.Targets.FirstOrDefault();
            if (!(firstTarget is ILocalScopeDeletionTarget localScopeTarget))
            {
                throw new ArgumentException();
            }

            return deletionGroup.PrecedingNonDeletedContext?.GetFollowingEndOfStatementContext()
                ?? (deletionGroup.OrderedFullDeletionTargets.LastOrDefault() == firstTarget || !firstTarget.IsFullDelete
                    ? localScopeTarget.ScopingContext.GetPrecedingEndOfStatementContext()
                    : localScopeTarget.ScopingContext.GetChild<VBAParser.EndOfStatementContext>());
        }
    }
}
