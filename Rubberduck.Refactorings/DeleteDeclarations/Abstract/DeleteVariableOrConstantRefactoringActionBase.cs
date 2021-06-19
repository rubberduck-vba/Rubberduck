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
    public abstract class DeleteVariableOrConstantRefactoringActionBase<TModel> : DeleteElementRefactoringActionBase<TModel> where TModel : class, IRefactoringModel
    {
        private static readonly string _lineContinuationExpression = $"{Tokens.LineContinuation}{Environment.NewLine}";

        public DeleteVariableOrConstantRefactoringActionBase(IDeclarationFinderProvider declarationFinderProvider, IDeclarationDeletionTargetFactory deletionTargetFactory, IDeleteDeclarationEndOfStatementContentModifierFactory eosModifierFactory, IRewritingManager rewritingManager)
            : base(declarationFinderProvider, deletionTargetFactory, eosModifierFactory, rewritingManager)
        {}

        protected override void ModifyDeletionGroups(IDeleteDeclarationsModel model, IModuleRewriter rewriter)
        {
            model.DeletionGroups.ForEach(dg => RemoveDeletionGroup(dg, model, rewriter));

            model.DeletionGroups.ForEach(dg => RemovePartialDeletions(dg, model.DeletionTargets, rewriter));
        }

        protected override List<IDeclarationDeletionTarget> CreateDeletionTargets(List<Declaration> targets, IDeclarationDeletionTargetFactory factory)
        {
            var deletionTargets = new List<IDeclarationDeletionTarget>();

            var remainingTargets = targets;

            while (remainingTargets.Any())
            {
                var deleteTarget = factory.Create(targets.First());

                if (deleteTarget.AllDeclarationsInListContext.Count >= 1)
                {
                    var listContextRelatedTargets = deleteTarget.AllDeclarationsInListContext.Intersect(targets);
                    deleteTarget.AddTargets(listContextRelatedTargets);
                    remainingTargets.RemoveAll(t => listContextRelatedTargets.Contains(t));
                }
                else
                {
                    remainingTargets.RemoveAll(t => t == targets.First());
                }
                deletionTargets.Add(deleteTarget);
            }

            return HandleLabelAndVarOrConstInSameBlock(deletionTargets);
        }

        protected virtual List<IDeclarationDeletionTarget> HandleLabelAndVarOrConstInSameBlock(List<IDeclarationDeletionTarget> blockDeleteTargets)
        {
            return blockDeleteTargets;
        }


        private static void RemovePartialDeletions(DeletionGroup deletionGroup, List<IDeclarationDeletionTarget> deleteTargets, IModuleRewriter rewriter)
        {
            foreach (var de in deletionGroup.Contexts)
            {
                var decDeleteTarget = deleteTargets.FirstOrDefault(d => d.TargetContext == de);
                if (decDeleteTarget?.IsFullDelete ?? true)
                {
                    continue;
                }

                //TODO: Need test(s) for Label with same-line multi variable list that is fully and partially deleted
                RemoveListDeclarationSubsetVariableOrConstant(decDeleteTarget, rewriter);
            }
        }
        private static void RemoveListDeclarationSubsetVariableOrConstant(IDeclarationDeletionTarget deleteDeclarationTarget, IModuleRewriter rewriter)
        {
            //Delete a subset of the the declaration list
            var retainedDeclarationsExpression = deleteDeclarationTarget.ListContext.GetText().Contains(_lineContinuationExpression)
                ? $"{BuildDeclarationsExpressionWithLineContinuations(deleteDeclarationTarget)}"
                : $"{string.Join(", ", deleteDeclarationTarget.RetainedDeclarations.Select(d => d.Context.GetText()))}";

            rewriter.Replace(deleteDeclarationTarget.ListContext.Parent, $"{GetDeclarationScopeExpression(deleteDeclarationTarget.TargetProxy)} {retainedDeclarationsExpression}");
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
