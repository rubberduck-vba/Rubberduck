using Antlr4.Runtime;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Rubberduck.Refactorings.DeleteDeclarations
{
    public struct DeletionGroup
    {
        public ParserRuleContext PrecedingNonDeletedContext { set; get; }

        public List<ParserRuleContext> Contexts { set; get; }

        public bool IsLastContext(ParserRuleContext parserRuleContext) => Contexts.Last() == parserRuleContext;
    }

    public abstract class DeleteElementRefactoringActionBase<TModel> : CodeOnlyRefactoringActionBase<TModel> where TModel : class, IRefactoringModel
    {
        protected readonly IDeclarationFinderProvider _declarationFinderProvider;
        protected readonly IRewritingManager _rewritingManager;
        private readonly IDeclarationDeletionTargetFactory _declarationDeletionTargetFactory;
        private readonly IDeleteDeclarationEndOfStatementContentModifierFactory _deleteDeclarationEndOfStatementContentModifierFactory;

        public DeleteElementRefactoringActionBase(IDeclarationFinderProvider declarationFinderProvider, IDeclarationDeletionTargetFactory deletionTargetFactory, IDeleteDeclarationEndOfStatementContentModifierFactory eosModifierFactory, IRewritingManager rewritingManager)
            : base(rewritingManager)
        {
            _declarationFinderProvider = declarationFinderProvider;
            _rewritingManager = rewritingManager;
            _declarationDeletionTargetFactory = deletionTargetFactory;
            _deleteDeclarationEndOfStatementContentModifierFactory = eosModifierFactory;
        }
        protected abstract void RefactorGuardClause(IDeleteDeclarationsModel model);

        protected abstract IOrderedEnumerable<ParserRuleContext> GetAllContextElements(Declaration declaration);

        protected Action<DeletionGroup, IDeclarationDeletionTarget> RetrieveNonDeleteDeclarationForGroup { private set;  get; } = (g, t) => g.PrecedingNonDeletedContext =  t.PrecedingEOSContext;

        protected void InjectRetrieveNonDeleteDeclarationForDeletionGroupAction(Action<DeletionGroup, IDeclarationDeletionTarget> setter)
            => RetrieveNonDeleteDeclarationForGroup = setter;

        protected void DeleteDeclarations(IDeleteDeclarationsModel model, IRewriteSession rewriteSession)
        {
            RefactorGuardClause(model);

            foreach (var targetGroup in model.Targets.ToLookup(t => t.QualifiedModuleName))
            {
                var deletionTargets = CreateDeletionTargets(model.Targets, _declarationDeletionTargetFactory);

                var deletionGroups = CreateDeletionGroups(deletionTargets, model.RemoveAllExceptionMessage);

                model.SetGroups(deletionGroups, deletionTargets);

                ModifyDeletionGroups(model, rewriteSession.CheckOutModuleRewriter(targetGroup.Key));
            }
        }

        protected virtual List<IDeclarationDeletionTarget> CreateDeletionTargets(List<Declaration> targets, IDeclarationDeletionTargetFactory factory)
            => factory.CreateMany(targets).ToList();

        protected IOrderedEnumerable<ParserRuleContext> GetAllTargetContextElements<TTarget, TDelete>(Declaration declaration)
            where TTarget : ParserRuleContext
            where TDelete : ParserRuleContext
        {
            var blockContext = declaration.Context.GetAncestor<TTarget>();

            return blockContext.children
                .Where(dt => dt is TDelete)
                .Cast<ParserRuleContext>()
                .OrderBy(c => c.GetSelection());
        }

       protected virtual void ModifyDeletionGroups(IDeleteDeclarationsModel model, IModuleRewriter rewriter)
            => model.DeletionGroups.ForEach(dg => RemoveDeletionGroup(dg, model, rewriter));

        protected void RemoveDeletionGroup(DeletionGroup deletionGroup, IDeleteDeclarationsModel model, IModuleRewriter rewriter)
        {
            foreach (var deleteContext in deletionGroup.Contexts)
            {
                //Remove the declaration
                var deleteTarget = model.DeletionTargets.FirstOrDefault(d => d.TargetContext == deleteContext);
                if (!deleteTarget?.IsFullDelete ?? true)
                {
                    //Removing a subset of declarations within a list is handled elsewhere
                    continue;
                }

                rewriter.Remove(deleteTarget.DeleteContext);
                
                if (deleteTarget.TargetProxy.Annotations.Any())
                {
                    //Note: deleting an Annotation deletes a context within the Preceding expression/declaration's EndOfStatementContext
                    foreach (var annotation in deleteTarget.TargetProxy.Annotations.Select(a => a.Context))
                    {
                        var annotationListIndividualEOFofEOSCtxt = annotation.GetAncestor<VBAParser.IndividualNonEOFEndOfStatementContext>();
                        rewriter.Remove(annotationListIndividualEOFofEOSCtxt);
                    }
                }

                if (deletionGroup.IsLastContext(deleteTarget.TargetContext))
                {
                    //Merge/Modify the EOSContext of the preceding non-deleted declaration with the EOSContext of the
                    //last declaration of the group
                    RetrieveNonDeleteDeclarationForGroup(deletionGroup, deleteTarget);
                    MergeEOSContexts(deleteTarget, model, rewriter);
                    break;
                }

                if (deleteTarget.EndOfStatementContext != null)
                {
                    rewriter.Remove(deleteTarget.EndOfStatementContext);
                }
            }
        }

        protected virtual void MergeEOSContexts(IDeclarationDeletionTarget deleteTarget, IDeleteDeclarationsModel model, IModuleRewriter rewriter)
        {
            var modifier = _deleteDeclarationEndOfStatementContentModifierFactory.Create();
            modifier.ModifyEndOfStatementContextContent(deleteTarget, model as IDeleteDeclarationModifyEndOfStatementContentModel, rewriter);
        }

        private static List<int> GetNonDeleteContextIndices(IOrderedEnumerable<ParserRuleContext> orderedContexts, IEnumerable<IDeclarationDeletionTarget> deleteDeclarationTargets)
        {
            var nonDeleteIndices = new List<int>();
            for (var idx = 0; idx < orderedContexts.Count(); idx++)
            {
                if (deleteDeclarationTargets.SingleOrDefault(dt => dt.TargetContext == orderedContexts.ElementAt(idx)) is null)
                {
                    nonDeleteIndices.Add(idx);
                }
            }
            return nonDeleteIndices;
        }

        private List<DeletionGroup> CreateDeletionGroups(List<IDeclarationDeletionTarget> deleteDeclarationTargets, string removeAllExceptionMessage = null)
        {
            var orderedContexts = GetAllContextElements(deleteDeclarationTargets.First().TargetProxy);

            var nonDeleteIndices = GetNonDeleteContextIndices(orderedContexts, deleteDeclarationTargets);

            if (removeAllExceptionMessage != null && !nonDeleteIndices.Any())
            {
                throw new InvalidOperationException(removeAllExceptionMessage);
            }

            return AssociatePrecedingNonDeleteContexts(nonDeleteIndices, orderedContexts, deleteDeclarationTargets);
        }

        private static List<DeletionGroup> AssociatePrecedingNonDeleteContexts(List<int> nonDeleteIndices, IEnumerable<ParserRuleContext> contexts, List<IDeclarationDeletionTarget> deleteDeclarationTargets)
        {
            if (!nonDeleteIndices.Any())
            {
                return new List<DeletionGroup>()
                {
                    new DeletionGroup()
                    {
                        PrecedingNonDeletedContext = null,
                        Contexts = contexts.ToList()
                    }
                };
            }

            var results = new List<DeletionGroup>();

            if (!nonDeleteIndices.Contains(0))
            {
                var deletionGroup = new DeletionGroup()
                {
                    PrecedingNonDeletedContext = null,
                    Contexts = contexts.Take(nonDeleteIndices.First()).ToList()
                };
                results.Add(deletionGroup);
            }

            for (var ndIdx = 0; ndIdx < nonDeleteIndices.Count; ndIdx++)
            {
                var firstNonDelete = nonDeleteIndices.ElementAt(ndIdx);

                var nextNonDelete = ndIdx + 1 < nonDeleteIndices.Count
                    ? contexts.ElementAt(nonDeleteIndices.ElementAt(ndIdx + 1))
                    : null;

                var toDelete = contexts.SkipWhile(ctxt => ctxt != contexts.ElementAt(firstNonDelete))
                    .Skip(1) //skip the leading nonDeleteContext
                    .TakeWhile(ctxt => ctxt != nextNonDelete);

                var deletionGroup = new DeletionGroup()
                {
                    PrecedingNonDeletedContext = contexts.ElementAt(ndIdx),
                    Contexts = toDelete.ToList()
                };
                results.Add(deletionGroup);
            }

            return results;
        }
    }
}
