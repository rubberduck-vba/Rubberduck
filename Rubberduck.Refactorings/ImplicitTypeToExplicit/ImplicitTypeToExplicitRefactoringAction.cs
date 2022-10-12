using System;
using System.Collections.Generic;
using System.Linq;
using Antlr4.Runtime;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.Refactorings.ImplicitTypeToExplicit
{
    public class ImplicitTypeToExplicitRefactoringAction : CodeOnlyRefactoringActionBase<ImplicitTypeToExplicitModel>
    {
        private readonly IDeclarationFinderProvider _declarationFinderProvider;
        private readonly IParseTreeValueFactory _parseTreeValueFactory;

        public ImplicitTypeToExplicitRefactoringAction(
            IDeclarationFinderProvider declarationFinderProvider,
            IParseTreeValueFactory parseTreeValueFactory,
            IRewritingManager rewritingManager)
            : base(rewritingManager)
        {
            _declarationFinderProvider = declarationFinderProvider;
            _parseTreeValueFactory = parseTreeValueFactory;
        }

        public override void Refactor(ImplicitTypeToExplicitModel model, IRewriteSession rewriteSession)
        {
            if (!(model.Target.Context is VBAParser.VariableSubStmtContext
                    || model.Target.Context is VBAParser.ConstSubStmtContext
                    || model.Target.Context is VBAParser.ArgContext))
            {
                throw new ArgumentException($"Invalid target {model.Target.IdentifierName}");
            }

            var identifierNode = model.Target.Context.GetChild<VBAParser.IdentifierContext>()
                        ?? model.Target.Context.GetChild<VBAParser.UnrestrictedIdentifierContext>() as ParserRuleContext;

            var insertAfterTarget = model.Target.IsArray
                ? model.Target.Context.Stop.TokenIndex
                : identifierNode.Stop.TokenIndex;

            var asTypeName = Tokens.Variant;
            if (!model.ForceVariantAsType)
            {
                var resolver = new ImplicitAsTypeNameResolver(_declarationFinderProvider, _parseTreeValueFactory, model.Target);
                asTypeName = InferAsTypeNameForInspectionResult(model.Target, resolver, new AsTypeNamesResultsHandler());
            }

            var rewriter = rewriteSession.CheckOutModuleRewriter(model.Target.QualifiedModuleName);
            rewriter.InsertAfter(insertAfterTarget, $" {Tokens.As} {asTypeName}");
        }

        private static string InferAsTypeNameForInspectionResult(Declaration target, ImplicitAsTypeNameResolver resolver, AsTypeNamesResultsHandler resultsHandler)
        {
            switch (target.DeclarationType)
            {
                case DeclarationType.Variable:
                    InferTypeNamesFromAssignmentLHSUsage(target, resolver, resultsHandler);
                    InferTypeNamesFromAssignmentRHSUsage(target, resolver, resultsHandler);
                    InferTypeNamesFromParameterUsage(target, resolver, resultsHandler);
                    break;

                case DeclarationType.Constant:
                    InferTypeNamesFromDeclarationWithDefaultValue(target.Context, resolver, resultsHandler);
                    InferTypeNamesFromAssignmentRHSUsage(target, resolver, resultsHandler);
                    InferTypeNamesFromParameterUsage(target, resolver, resultsHandler);
                    break;

                case DeclarationType.Parameter:
                    if (target.Context.TryGetChildContext<VBAParser.ArgDefaultValueContext>(out var argDefaultValueCtxt))
                    {
                        InferTypeNamesFromDeclarationWithDefaultValue(argDefaultValueCtxt, resolver, resultsHandler);
                    }

                    InferTypeNamesFromAssignmentLHSUsage(target, resolver, resultsHandler);
                    InferTypeNamesFromAssignmentRHSUsage(target, resolver, resultsHandler);
                    InferTypeNamesFromParameterUsage(target, resolver, resultsHandler);
                    break;
            }

            return resultsHandler.ResolveAsTypeName(target);
        }

        private static void InferTypeNamesFromParameterUsage(Declaration target, ImplicitAsTypeNameResolver resolver, AsTypeNamesResultsHandler resultsHandler)
        {
            var argumentListContexts = target.References
                .Select(rf => rf.Context.GetAncestor<VBAParser.ArgumentListContext>())
                .Where(c => c != null);

            if (argumentListContexts.Any())
            {
                resultsHandler.AddCandidates(nameof(VBAParser.ArgumentListContext), resolver.InferAsTypeNames(argumentListContexts));
            }
        }

        private static void InferTypeNamesFromDeclarationWithDefaultValue(ParserRuleContext context, ImplicitAsTypeNameResolver resolver, AsTypeNamesResultsHandler resultsHandler)
        {
            var results = new Dictionary<string, List<string>>();

            var lExprContext = context.GetChild<VBAParser.LExprContext>();
            var litExprContext = context.GetChild<VBAParser.LiteralExprContext>();
            var concatExprContext = context.GetChild<VBAParser.ConcatOpContext>();

            if (lExprContext is null && litExprContext is null && concatExprContext is null)
            {
                //Declarations that have a default value expression (Constants and Optional parameters) 
                //must resolve to an AsTypeName. Expressions are indeterminant and result assigning the Variant type
                resultsHandler.AddIndeterminantResult();
                return;
            }

            if (lExprContext != null)
            {
                resultsHandler.AddCandidates(nameof(VBAParser.LExprContext), resolver.InferAsTypeNames(new List<VBAParser.LExprContext>() { lExprContext }));
            }

            if (litExprContext != null)
            {
                resultsHandler.AddCandidates(nameof(VBAParser.LiteralExprContext), resolver.InferAsTypeNames(new List<VBAParser.LiteralExprContext>() { litExprContext }));
            }

            if (concatExprContext != null)
            {
                resultsHandler.AddCandidates(nameof(VBAParser.ConcatOpContext), resolver.InferAsTypeNames(new List<VBAParser.ConcatOpContext>() { concatExprContext }));
            }
        }

        private static void InferTypeNamesFromAssignmentLHSUsage(Declaration target, ImplicitAsTypeNameResolver resolver, AsTypeNamesResultsHandler resultsHandler)
        {
            var assignmentContextsToEvaluate = target.References
                .Where(rf => rf.IsAssignment && rf.Context.Parent != null)
                .Select(rf => rf.Context.Parent).Cast<ParserRuleContext>();

            if (!assignmentContextsToEvaluate.Any())
            {
                return;
            }

            List<T> AssignmentRHSContexts<T>() where T: ParserRuleContext
                => assignmentContextsToEvaluate.Select(c => c.GetChild<T>())
                    .Where(c => c != null).ToList();

            //Until a unified Expression engine is available, the following are the only ParserRuleContext
            //Subclasses that are evaluated.
            var newExprContexts = AssignmentRHSContexts<VBAParser.NewExprContext>().ToList();
            var lExprContexts = AssignmentRHSContexts<VBAParser.LExprContext>().ToList();
            var litExprContexts = AssignmentRHSContexts<VBAParser.LiteralExprContext>().ToList();
            var concatExprContexts = AssignmentRHSContexts<VBAParser.ConcatOpContext>().ToList();

            var countOfAllContexts = SumContextCounts<VBAParser.ExpressionContext>(
                newExprContexts, lExprContexts, litExprContexts, concatExprContexts);

            if (assignmentContextsToEvaluate.Count() == countOfAllContexts)
            {
                resultsHandler.AddCandidates(nameof(VBAParser.NewExprContext), resolver.InferAsTypeNames(newExprContexts));
                resultsHandler.AddCandidates(nameof(VBAParser.LExprContext), resolver.InferAsTypeNames(lExprContexts));
                resultsHandler.AddCandidates(nameof(VBAParser.LiteralExprContext), resolver.InferAsTypeNames(litExprContexts));
                resultsHandler.AddCandidates(nameof(VBAParser.ConcatOpContext), resolver.InferAsTypeNames(concatExprContexts));
                return;
            }

            resultsHandler.AddIndeterminantResult();
        }

        private static long SumContextCounts<T>(params IEnumerable<T>[] contextLists) where T : VBAParser.ExpressionContext
            => contextLists.Sum(c => c.Count());

        private static void InferTypeNamesFromAssignmentRHSUsage(Declaration target, ImplicitAsTypeNameResolver resolver, AsTypeNamesResultsHandler resultsHandler)
        {
            var rhsLetStmtContexts = target.References
                .Where(rf => !rf.IsAssignment
                    && rf.Context.Parent is VBAParser.LExprContext lExpr
                    && lExpr.Parent is VBAParser.LetStmtContext)
                .Select(rf => rf.Context.GetAncestor<VBAParser.LetStmtContext>())
                .ToList();

            if (rhsLetStmtContexts.Any())
            {
                resultsHandler.AddCandidates(nameof(VBAParser.LetStmtContext), resolver.InferAsTypeNames(rhsLetStmtContexts));
            }
        }

    }
}
