using Antlr4.Runtime;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Rubberduck.Refactorings.ImplicitTypeToExplicit
{
    /// <summary>
    /// ConcatOpContextResolver resolves the AsTypeName of ConcatOpContext expressions 
    /// assumed to be on the RHS of a Variable, Constant, or Parameter assignment.  
    /// </summary>
    public class ConcatOpContextResolver
    {
        private readonly IDeclarationFinderProvider _declarationFinderProvider;

        public ConcatOpContextResolver(IDeclarationFinderProvider declarationFinderProvider)
        {
            _declarationFinderProvider = declarationFinderProvider;
        }
        /// <summary>
        /// Returns an AsTypeName result of 'String' or 'Variant' for a List of 
        /// ConcatOpContext expressions.
        /// </summary>
        /// <remarks>
        /// Until a unified Expression engine is available, the default AsTypeName, 
        /// 'Variant', is returned for all ConcatOpContexts that contain operand 
        /// contexts other than LiteralExprContexts and LExprContexts.  
        /// 
        /// In general, the '&amp;' operator (and this function) returns 'String' 
        /// unless a 'Variant' operand is found within the expression.
        /// </remarks>
        public List<string> InferAsTypeNames(IEnumerable<VBAParser.ConcatOpContext> tContexts)
        {
            if (!tContexts.Any())
            {
                return new List<string>();
            }

            var operandContexts = GetConcatOperandContexts(tContexts).ToList();

            //The logic below will incorrectly interpret a (very unlikely) statement 
            //like '5 & Null & Null & 5' as a 'Variant' instead of the correct 
            //AsTypename 'String' ("55").
            //TODO: The issue above can be resolved once a unified Expression engine 
            //is available,
            if (tContexts.Any(ctxt => ctxt.GetText().Contains("Null & Null")))
            {
                return new List<string>() { Tokens.Variant };
            }

            var literals = operandContexts.OfType<VBAParser.LiteralExprContext>();
            var lexprs = operandContexts.OfType<VBAParser.LExprContext>();
            if (operandContexts.Count() != (literals.Count() + lexprs.Count()))
            {
                //A context type other than VBAParser.LiteralExprContext 
                //or VBAParser.LExprContext is used - resort to the default AsTypeName
                return new List<string>() { Tokens.Variant };
            }

            if (InferAsTypeNamesForLExprContexts(lexprs, _declarationFinderProvider)
                .Any(tn => tn == Tokens.Variant))
            {
                return new List<string>() { Tokens.Variant };
            }

            return new List<string>() { Tokens.String };
        }

        private static IEnumerable<ParserRuleContext> GetConcatOperandContexts(
            IEnumerable<VBAParser.ConcatOpContext> tContexts)
        {
            var results = new List<ParserRuleContext>();
            foreach (var ctxt in tContexts)
            {
                results = ExtractOperands(ctxt, results);
            }

            return results;
        }

        private static List<ParserRuleContext> ExtractOperands(
            VBAParser.ConcatOpContext concatOpContext,
            List<ParserRuleContext> operandContexts)
        {
            if (concatOpContext.children.First() is VBAParser.ConcatOpContext concatOpCtxt)
            {
                operandContexts = ExtractOperands(concatOpCtxt, operandContexts);
            }

            var operands = new List<ParserRuleContext>
                { concatOpContext.children.First() as ParserRuleContext,
                concatOpContext.children.Last() as ParserRuleContext};

            foreach (var operandContext in operands)
            {
                if (!(operandContext is VBAParser.ConcatOpContext))
                {
                    operandContexts.Add(operandContext);
                }
            }

            return operandContexts;
        }
        private static List<string> InferAsTypeNamesForLExprContexts(
            IEnumerable<VBAParser.LExprContext> lExprCtxts, 
            IDeclarationFinderProvider declarationFinderProvider)
        {
            var results = new List<string>();

            string typeNameResult;
            foreach (var lExpr in lExprCtxts)
            {
                var target = GetLExprDeclaration(lExpr, declarationFinderProvider);
                typeNameResult = target.IsObject ? Tokens.Variant : target.AsTypeName;
                results.Add(typeNameResult);
            }
            return results;
        }

        private static Declaration GetLExprDeclaration(
            VBAParser.LExprContext lExprContext, 
            IDeclarationFinderProvider declarationFinderProvider)
        {
            var simpleNameExpression = 
                lExprContext.GetDescendent<VBAParser.SimpleNameExprContext>();
            
            var candidates = declarationFinderProvider.DeclarationFinder
                .MatchName(simpleNameExpression.GetText());

            if (candidates.Count() == 1)
            {
                return candidates.First();
            }

            var lExprDeclaration = candidates
                .Single(c => c.References.Any(rf => rf.Context is ParserRuleContext prc
                    && simpleNameExpression.Equals(prc)));

            return lExprDeclaration;
        }
    }
}
