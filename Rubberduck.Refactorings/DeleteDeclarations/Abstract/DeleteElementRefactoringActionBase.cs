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
    }

    public abstract class DeleteElementRefactoringActionBase<TModel> : CodeOnlyRefactoringActionBase<TModel> where TModel : class, IRefactoringModel
    {
        protected readonly IDeclarationFinderProvider _declarationFinderProvider;
        protected readonly IRewritingManager _rewritingManager;
        
        private static readonly string _lineContinuationExpression = $"{Tokens.LineContinuation}{Environment.NewLine}";

        public DeleteElementRefactoringActionBase(IDeclarationFinderProvider declarationFinderProvider, IRewritingManager rewritingManager)
            : base(rewritingManager)
        {
            _declarationFinderProvider = declarationFinderProvider;
            _rewritingManager = rewritingManager;
        }

        protected void DeleteTargetsInModuleImpl<TList, TTarget>(List<IDeclarationDeletionTarget> deleteTargets, IModuleRewriter rewriter)
            where TList : ParserRuleContext
            where TTarget : ParserRuleContext
        {
            var contextElements = GetAllContextElements<TList, TTarget>(deleteTargets.First().TargetProxy);

            var deletionGroups = CreateDeletionGroups(deleteTargets, contextElements);

            foreach (var deletionGroup in deletionGroups)
            {
                DeleteGroup(deletionGroup, deleteTargets, rewriter);
            }
        }
        protected virtual List<DeletionGroup> CreateDeletionGroups(List<IDeclarationDeletionTarget> deleteDeclarationTargets, IOrderedEnumerable<ParserRuleContext> allModuleContextElements)
        {
            var nonDeleteIndices = GetNonDeleteContextIndices(allModuleContextElements, deleteDeclarationTargets);

            return AssociatePrecedingNonDeleteContexts(nonDeleteIndices, allModuleContextElements, deleteDeclarationTargets);
        }

        protected virtual List<int> GetNonDeleteContextIndices(IOrderedEnumerable<ParserRuleContext> orderedContexts, IEnumerable<IDeclarationDeletionTarget> deleteDeclarationTargets)
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

        protected List<DeletionGroup> AssociatePrecedingNonDeleteContexts(List<int> nonDeleteIndices, IEnumerable<ParserRuleContext> contexts, List<IDeclarationDeletionTarget> deleteDeclarationTargets)
        {
            var contextToDeleteTarget = new Dictionary<ParserRuleContext, IDeclarationDeletionTarget>();
            foreach (var ctxt in contexts)
            {
                var dTarget = deleteDeclarationTargets.SingleOrDefault(dt => dt.TargetContext == ctxt);
                if (dTarget != null)
                {
                    contextToDeleteTarget.Add(ctxt, dTarget);
                }
            }

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
        protected virtual IOrderedEnumerable<ParserRuleContext> GetAllContextElements<TTarget, TDelete>(Declaration declaration)
            where TTarget : ParserRuleContext
            where TDelete : ParserRuleContext
        {
            var blockContext = declaration.Context.GetAncestor<TTarget>();

            return blockContext.children
                .Where(dt => dt is TDelete)
                .Cast<ParserRuleContext>()
                .OrderBy(c => c.GetSelection());
        }

        protected void DeleteGroup(DeletionGroup deletionGroup, List<IDeclarationDeletionTarget> deleteTargets, IModuleRewriter rewriter)
        {
            foreach (var deleteContext in deletionGroup.Contexts)
            {
                var deleteTarget = deleteTargets.FirstOrDefault(d => d.TargetContext == deleteContext);
                if (!deleteTarget?.IsFullDelete ?? true)
                {
                    continue;
                }

                rewriter.Remove(deleteTarget.DeleteContext);

                RemoveAnnotations(deleteTarget.TargetProxy, rewriter);

                ModifyEOSContexts(deletionGroup, deleteTarget, rewriter);
            }
        }

        protected virtual void ModifyEOSContexts(DeletionGroup deletionGroup, IDeclarationDeletionTarget deleteTarget, IModuleRewriter rewriter)
        {
            if (deletionGroup.PrecedingNonDeletedContext is null && deleteTarget.EndOfStatementContext is null)
            {
                return;
            }

            var targetEOSContextContentProvider = new EOSContextContentProvider(deleteTarget.EndOfStatementContext, rewriter);
            if (targetEOSContextContentProvider.HasDeclarationLogicalLineComment)
            {
                rewriter.Remove(targetEOSContextContentProvider.DeclarationLogicalLineCommentContext.GetChild(0));
            }

            if (deletionGroup.Contexts.Last() != deleteTarget.TargetContext)
            {
                rewriter.Remove(deleteTarget.EndOfStatementContext);
                return;
            }
            ModifyEOSContextForLastTargetOfDeletionGroup(deleteTarget, rewriter);
        }

        /// <summary>
        /// Modify the proceding non-deleted EOS with EOS content from the last deleted target.
        /// </summary>
        protected virtual void ModifyEOSContextForLastTargetOfDeletionGroup(IDeclarationDeletionTarget deleteTarget, IModuleRewriter rewriter)
        {
            var groupEndingEOSContentProvider = new EOSContextContentProvider(deleteTarget.EndOfStatementContext, rewriter);

            if (groupEndingEOSContentProvider.ModifiedEOSContent.StartsWith(": "))
            {
                rewriter.Remove(deleteTarget.EndOfStatementContext);
                return;
            }

            var precedingEOSContentProvider = new EOSContextContentProvider(deleteTarget.PrecedingEOSContext, rewriter);

            string replacementContent;

            if (groupEndingEOSContentProvider.ModifiedContentContainsCommentMarker)
            {
                replacementContent = precedingEOSContentProvider.ModifiedContentContainsCommentMarker
                    ? $"{precedingEOSContentProvider.ContentPriorToSeparationAndIndentation}{groupEndingEOSContentProvider.ModifiedEOSContent}"
                    : $"{precedingEOSContentProvider.ContentPriorToSeparationAndIndentation}{precedingEOSContentProvider.Separation}{groupEndingEOSContentProvider.ContentFreeOfStartingNewLines}";
            }
            else
            {
                replacementContent = $"{precedingEOSContentProvider.ContentPriorToSeparationAndIndentation}{precedingEOSContentProvider.Separation}{groupEndingEOSContentProvider.Indentation}";
                if (deleteTarget.PrecedingEOSContext.TryGetPrecedingContext<VBAParser.ModuleBodyElementContext>(out var mbe))
                {
                    var minSeparationForProceduresAfterDeletions = string.Concat(Enumerable.Repeat(Environment.NewLine, 2));
                    if (!precedingEOSContentProvider.Separation.Contains(minSeparationForProceduresAfterDeletions))
                    {
                        replacementContent = $"{precedingEOSContentProvider.ContentPriorToSeparationAndIndentation}{precedingEOSContentProvider.Separation}{Environment.NewLine}{groupEndingEOSContentProvider.Indentation}";
                    }
                }
            }

            //Replace the precedingEOSContent if it exists, or modify the last deleted target's EOSContext
            if (deleteTarget.PrecedingEOSContext != null)
            {
                rewriter.Replace(deleteTarget.PrecedingEOSContext, replacementContent);
                if (deleteTarget.EndOfStatementContext != null)
                {
                     rewriter.Remove(deleteTarget.EndOfStatementContext);
                }
            }
            else if (deleteTarget.EndOfStatementContext != null)
            {
                rewriter.Replace(deleteTarget.EndOfStatementContext, replacementContent);
            }
        }

        protected static void RemoveAnnotations(Declaration declaration, IModuleRewriter rewriter)
        {
            foreach (var annotation in declaration.Annotations.Select(a => a.Context))
            {
                var annotationListIndividualEOFofEOSCtxt = annotation.GetAncestor<VBAParser.IndividualNonEOFEndOfStatementContext>();

                //Note: deleting an Annotation deletes a context within the Preceding expression/declaration's EndOfStatementContext
                rewriter.Remove(annotationListIndividualEOFofEOSCtxt);
            }
        }

        protected static void RemoveListDeclarationSubsetVariableOrConstant(IDeclarationDeletionTarget deleteDeclarationTarget, IModuleRewriter rewriter)
        {
            //Delete a subset of the the declaration list
            var retainedDeclarationsExpression = deleteDeclarationTarget.ListContext.GetText().Contains(_lineContinuationExpression)
                ? $"{BuildDeclarationsExpressionWithLineContinuations(deleteDeclarationTarget)}"
                : $"{string.Join(", ", deleteDeclarationTarget.RetainedDeclarations.Select(d => d.Context.GetText()))}";

            rewriter.Replace(deleteDeclarationTarget.ListContext.Parent, $"{GetDeclarationScopeExpression(deleteDeclarationTarget.TargetProxy)} {retainedDeclarationsExpression}");
        }

        private static string GetDeclarationScopeExpression(Declaration listPrototype)
        {
            if (listPrototype.DeclarationType.HasFlag(DeclarationType.Variable))
            {
                var accessToken = listPrototype.Accessibility == Accessibility.Implicit
                    ? Tokens.Private
                    : $"{listPrototype.Accessibility}";

                return listPrototype.ParentDeclaration is ModuleDeclaration
                    ? accessToken
                    : Tokens.Dim;
            }

            if (listPrototype.DeclarationType.HasFlag(DeclarationType.Constant))
            {
                var accessToken = listPrototype.Accessibility == Accessibility.Implicit
                    ? Tokens.Private
                    : $"{listPrototype.Accessibility}";

                return listPrototype.ParentDeclaration is ModuleDeclaration
                    ? $"{accessToken} {Tokens.Const}"
                    : Tokens.Const;
            }

            throw new ArgumentException("Unsupported DeclarationType");
        }

        private static string BuildDeclarationsExpressionWithLineContinuations(IDeclarationDeletionTarget deleteDeclarationTarget)
        {
            var elementsByLineContinuation = deleteDeclarationTarget.ListContext.GetText().Split(new string[] { _lineContinuationExpression }, StringSplitOptions.None);

            if (elementsByLineContinuation.Count() == 1)
            {
                throw new ArgumentException("'targetsToDelete' parameter does not contain line extension(s)");
            }

            var expr = new StringBuilder();
            foreach (var element in elementsByLineContinuation)
            {
                var idContexts = deleteDeclarationTarget.RetainedDeclarations.Where(r => element.Contains(r.Context.GetText())).Select(d => d);
                foreach (var ctxt in idContexts)
                {
                    var indent = string.Concat(element.TakeWhile(e => e == ' '));

                    expr = expr.Length == 0
                        ? expr.Append(ctxt.Context.GetText())
                        : expr.Append($",{_lineContinuationExpression}{indent}{ctxt.Context.GetText()}");
                }
            }
            return expr.ToString();
        }

        //TODO: Used for debugging - delete eventually
        protected static string GetModifiedContextText(ParserRuleContext prContext, IModuleRewriter rewriter)
            => rewriter.GetText(prContext.Start.TokenIndex, prContext.Stop.TokenIndex);
        //TODO: Keep this for reference
        private static bool LabelIsOnlyContentInBlockStatementContext(ParserRuleContext context, out VBAParser.BlockStmtContext blockStmt)
        {
            return context.TryGetAncestor(out blockStmt)
                && blockStmt.ChildCount == 1
                && blockStmt.children.First() == context.Parent;//Parent is VBAParser.StatementLabelDefinitionContext
        }
    }
}
