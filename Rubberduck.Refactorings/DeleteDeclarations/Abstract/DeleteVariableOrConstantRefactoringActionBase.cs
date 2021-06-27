using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Rubberduck.Refactorings.DeleteDeclarations.Abstract
{
    public abstract class DeleteVariableOrConstantRefactoringActionBase<TModel> : DeleteElementRefactoringActionBase<TModel> where TModel : class, IRefactoringModel
    {
        private static readonly string _lineContinuationExpression = $"{Tokens.LineContinuation}{Environment.NewLine}";

        public DeleteVariableOrConstantRefactoringActionBase(IDeclarationFinderProvider declarationFinderProvider, 
            IDeclarationDeletionTargetFactory deletionTargetFactory,
            IDeclarationDeletionGroupsGeneratorFactory deletionGroupsGeneratorFactory,
            IRewritingManager rewritingManager)
            : base(declarationFinderProvider, deletionTargetFactory, deletionGroupsGeneratorFactory,rewritingManager)
        {}

        protected override void RemoveDeletionGroups(IEnumerable<IDeclarationDeletionGroup> deletionGroups, IDeleteDeclarationsModel model, IModuleRewriter rewriter)
        {
            foreach (var deletionGroup in deletionGroups)
            {
                if (deletionGroup.OrderedPartialDeletionTargets.Any())
                {
                    RemovePartialDeletionTargets(deletionGroup, model, rewriter);
                }

                if (deletionGroup.OrderedFullDeletionTargets.Any())
                {
                    RemoveFullDeletionGroup(deletionGroup, model, rewriter);
                }
            }
        }

        private void RemovePartialDeletionTargets(IDeclarationDeletionGroup deletionGroup, IDeleteDeclarationsModel model, IModuleRewriter rewriter)
        {
            var lastTarget = deletionGroup.OrderedPartialDeletionTargets.Last();

            lastTarget.PrecedingEOSContext = GetPrecedingNonDeletedEOSContextForGroup(deletionGroup);

            var retainedDeclarationsExpression = lastTarget.ListContext.GetText().Contains(_lineContinuationExpression)
                ? $"{BuildDeclarationsExpressionWithLineContinuations(lastTarget)}"
                : $"{string.Join(", ", lastTarget.RetainedDeclarations.Select(d => d.Context.GetText()))}";

            rewriter.Replace(lastTarget.ListContext.Parent, $"{GetDeclarationScopeExpression(lastTarget.TargetProxy)} {retainedDeclarationsExpression}");

            if (lastTarget is null || lastTarget.TargetEOSContext is null)
            {
                return;
            }

            if (lastTarget.TargetEOSContext.GetText() == EOS_COLON)
            {
                //Remove the declarations EOS colon character and use the PrecedingEOSContext as-is
                lastTarget.Rewriter.Remove(lastTarget.TargetEOSContext);
                return;
            }

            ModifyRelatedComments(lastTarget, model, rewriter);

            rewriter.Replace(lastTarget.TargetEOSContext, lastTarget.ModifiedTargetEOSContent);
        }

        protected override IEnumerable<IDeclarationDeletionTarget> CreateDeletionTargets(IEnumerable<Declaration> declarations, IRewriteSession rewriteSession, IDeclarationDeletionTargetFactory targetFactory)
        {
            var deletionTargets = new List<IDeclarationDeletionTarget>();

            var remainingTargets = declarations.ToList();

            while (remainingTargets.Any())
            {
                var deleteTarget = targetFactory.Create(remainingTargets.First(), rewriteSession);

                if (deleteTarget.AllDeclarationsInListContext.Count >= 1)
                {
                    var listContextRelatedTargets = deleteTarget.AllDeclarationsInListContext.Intersect(declarations);
                    deleteTarget.AddTargets(listContextRelatedTargets);
                    remainingTargets.RemoveAll(t => listContextRelatedTargets.Contains(t));
                }
                else
                {
                    remainingTargets.RemoveAll(t => t == declarations.First());
                }

                deletionTargets.Add(deleteTarget);
            }

            return deletionTargets;
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
    }
}