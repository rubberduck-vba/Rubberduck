using Antlr4.Runtime;
using Antlr4.Runtime.Misc;
using Antlr4.Runtime.Tree;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Results;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Inspections.Resources;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using static Rubberduck.Parsing.Grammar.VBAParser;

namespace Rubberduck.Inspections.Concrete
{
    public interface ISelectStmtClauseVisitor
    {
        void Visit(VBAParser.SelectCaseStmtContext context);
        void Visit(VBAParser.SelectExpressionContext context);
        void Visit(VBAParser.CaseClauseContext context);
        void Visit(VBAParser.RangeClauseContext context);
        void Visit(VBAParser.SelectStartValueContext context);
        void Visit(VBAParser.SelectEndValueContext context);
        void Visit(VBAParser.RelationalOpContext context);
        void Visit(VBAParser.MultOpContext context);
        void Visit(VBAParser.AddOpContext context);
        void Visit(VBAParser.PowOpContext context);
        void Visit(VBAParser.ModOpContext context);
        void Visit(VBAParser.UnaryMinusOpContext context);
        void Visit(VBAParser.LogicalAndOpContext context);
        void Visit(VBAParser.LogicalOrOpContext context);
        void Visit(VBAParser.LogicalXorOpContext context);
        void Visit(VBAParser.LogicalEqvOpContext context);
        //void Visit(VBAParser.LogicalImpOpContext context);
        void Visit(VBAParser.LogicalNotOpContext context);
        void Visit(VBAParser.ParenthesizedExprContext context);
        void Visit(VBAParser.LExprContext context);
        void Visit(VBAParser.LiteralExprContext context);
    }

    public interface ISelectStmtClause
    {
        void Accept(ISelectStmtClauseVisitor visitor);
        ParserRuleContext Context { get; }
    }

    public interface ISupportTestsUnreachableCaseInspection
    {
        List<ParserRuleContext> CaseClauseContextsForSelectStmt(ParserRuleContext selectStmt);
    }

    public sealed class UnreachableCaseInspection : ParseTreeInspectionBase, ISupportTestsUnreachableCaseInspection
    {
        public Dictionary<ParserRuleContext, List<ParserRuleContext>> BuildSelectToCasesHierarchy(IEnumerable<QualifiedContext<ParserRuleContext>> selectStmts)
        {
            var result = new Dictionary<ParserRuleContext, List<ParserRuleContext>>();
            foreach (var selectStmt in selectStmts)
            {
                result.Add(selectStmt.Context, selectStmt.Context.GetChildren<CaseClauseContext>().Select(cc => (ParserRuleContext)cc).ToList());
            }
            return result;
        }

        public List<ParserRuleContext> CaseClauseContextsForSelectStmt(ParserRuleContext selectStmt)
        {
            return selectStmt.GetChildren<CaseClauseContext>().Select(cc => (ParserRuleContext)cc).ToList();
        }

        public List<ParserRuleContext> RangeContextsForCase(ParserRuleContext caseCtxt)
        {
            return caseCtxt.GetChildren<RangeClauseContext>().Select(cc => (ParserRuleContext)cc).ToList();
        }
        //public enum ClauseEvaluationResult { Unreachable, MismatchType, CaseElse, NoResult };

        //internal Dictionary<ClauseEvaluationResult, string> ResultMessages = new Dictionary<ClauseEvaluationResult, string>()
        //{
        //    [ClauseEvaluationResult.Unreachable] = InspectionsUI.UnreachableCaseInspection_Unreachable,
        //    [ClauseEvaluationResult.MismatchType] = InspectionsUI.UnreachableCaseInspection_TypeMismatch,
        //    [ClauseEvaluationResult.CaseElse] = InspectionsUI.UnreachableCaseInspection_CaseElse
        //};

        //Used to modify logic operators to inspect expressions like '5 > x' as 'x < 5'
        private static Dictionary<string, string> AlgebraicLogicalInversions = new Dictionary<string, string>()
        {
            [CompareTokens.EQ] = CompareTokens.EQ,
            [CompareTokens.NEQ] = CompareTokens.NEQ,
            [CompareTokens.LT] = CompareTokens.GT,
            [CompareTokens.LTE] = CompareTokens.GTE,
            [CompareTokens.GT] = CompareTokens.LT,
            [CompareTokens.GTE] = CompareTokens.LTE
        };

        public UnreachableCaseInspection(RubberduckParserState state)
            : base(state, CodeInspectionSeverity.Suggestion) { }

        public override IInspectionListener Listener { get; } =
            new UnreachableCaseInspectionListener();

        public override CodeInspectionType InspectionType => CodeInspectionType.CodeQualityIssues;

        protected override IEnumerable<IInspectionResult> DoGetInspectionResults()
        {
            var inspResults = new List<IInspectionResult>();

            var selectCaseContexts = Listener.Contexts
                .Where(result => !IsIgnoringInspectionResultFor(result.ModuleName, result.Context.Start.Line));

            var exprValues = new Dictionary<ParserRuleContext, UnreachableCaseInspectionValue>();
            foreach (var selectStmt in selectCaseContexts)
            {
            //    Debug.Assert(selectStmt.Context.GetDescendent<VBAParser.SelectExpressionContext>() != null);
            //    var selectExprCtxt = selectStmt.Context.GetDescendent<VBAParser.SelectExpressionContext>();

            //    var lExprs = selectExprCtxt.GetDescendents<LExprContext>();
            //    foreach (var expr in lExprs)
            //    {
            //        if (!exprValues.ContainsKey(expr))
            //        {
            //            exprValues.Add(expr, CreateValue(expr));
            //        }
            //    }
            //    var result = exprValues.All(expr => expr.Value.UseageTypeName.Equals(exprValues.Values.First().UseageTypeName));
            //    var allLExprs = selectStmt.Context.GetDescendents<LExprContext>();
            //    foreach (var expr in allLExprs)
            //    {
            //        if (!exprValues.ContainsKey(expr))
            //        {
            //            exprValues.Add(expr, CreateValue(expr));
            //        }
            //    }
            }
            //var selectStmtClauses = new List<ISelectStmtClause>();
            //foreach (var selectStmt in selectCaseContexts)
            //{
            //    var test = new SelectContext(selectStmt.Context);
            //    selectStmtClauses.Add(test);
            //}

            //foreach (var ssClause in selectStmtClauses)
            //{
            //    //var test = ssClause.StringCheck;
            //    ssClause.Accept(new SummaryCoverage(State, exprValues));
            //    //var test2 = ssClause.StringCheck;
            //}
            return inspResults;
        }

        public SummaryCoverage GetCoverage(VBAParser.RangeClauseContext rgClause, string EvaluationTypeName, SummaryCoverage test = null)
        {
            if (test == null)
            {
                test = new SummaryCoverage(State);
            }
            test.EvaluationTypeName = EvaluationTypeName;

            var rg = new SelectContext(rgClause);
            rg.Accept(test);
            return test;
        }

        public List<UnreachableCaseInspectionValue> GetCoverage2(VBAParser.SelectCaseStmtContext selectCaseStmtCtxt, string EvaluationTypeName, SummaryCoverage test = null)
        {
            if (test == null)
            {
                test = new SummaryCoverage(State);
            }
            test.EvaluationTypeName = EvaluationTypeName;

            var exprs = selectCaseStmtCtxt.GetDescendents().Where(desc => desc is VBAParser.LExprContext).ToList();// || desc is LiteralExprContext).ToList();
            var values = new List<UnreachableCaseInspectionValue>();
            for(var idx = 0; idx < exprs.Count(); idx++)
            {
                var rg = new SelectContext((ParserRuleContext)exprs[idx]);
                rg.Accept(test);
                values.AddRange(test.VariableCtxts.Values);
                values.AddRange(test.ConstantCtxts.Values);
            }


            return values;
        }

        public string GetSelectCaseEvaluationType(VBAParser.SelectCaseStmtContext selectStmt)
        {
            var exprValues = new List<UnreachableCaseInspectionValue>();
            var selectExprCtxt = selectStmt.GetChild<VBAParser.SelectExpressionContext>();

            if (selectExprCtxt.ChildCount == 1 && selectExprCtxt.children
                .Any(child => IsLogicalContext(child) || EqualsTrueFalseLiteral(child)))
            {
                return Tokens.Boolean;
            }

            if (selectExprCtxt.ChildCount == 1)
            {
                var smplName = selectExprCtxt.GetDescendent<VBAParser.SimpleNameExprContext>();
                if (SymbolList.TypeHintToTypeName.TryGetValue(smplName.GetText().Last().ToString(), out string theTypeName))
                {
                    return theTypeName;
                }
            }

            var values = GetCoverage2(selectStmt, "");
            var typeNames = values.Select(val => val.UseageTypeName).ToList();

            var typeName = typeNames.Any() ? DetermineEvaluationTypeFromTypes(typeNames) : string.Empty;
            if (typeName.Equals(Tokens.Variant) || typeName.Equals(string.Empty))
            {
                var rgClauses = selectStmt.GetDescendents<VBAParser.RangeClauseContext>();
                var rgClauseTypeResult = GetSelectCaseEvaluationType(rgClauses);
                return rgClauseTypeResult;
            }
            return typeName;
        }

        public string GetSelectCaseEvaluationType(IEnumerable<VBAParser.RangeClauseContext> rgClauses)
        {
            var exprValues = new List<UnreachableCaseInspectionValue>();
            var exprs = rgClauses.SelectMany(rg => rg.GetDescendents())
                .Where(expr => expr is VBAParser.LExprContext || expr is VBAParser.LiteralExprContext);
            foreach (var expr in exprs)
            {
                var test = new SummaryCoverage(State);
                var rg = new SelectContext((ParserRuleContext)expr);
                rg.Accept(test);
                exprValues.AddRange(test.VariableCtxts.Values);
                exprValues.AddRange(test.ConstantCtxts.Values);
            }
            var typeNames = exprValues.Select(expr => expr.UseageTypeName);
            var typeName = DetermineEvaluationTypeFromTypes(typeNames);

            //If Strings are in the mix and prevent resolution to a type, we remove them
            //here and see if a resolution becomes possible.  The strings will be converted to the
            //final type during subsequent unreachable analysis.  If they cannot be converted to
            //the "Evaluation Type", they will be flagged as mismatching e.g., "45" converts to a number
            //but "foo" will not.
            if (typeName.Equals(string.Empty))
            {
                var modifiedNames = typeNames.ToList();
                modifiedNames.RemoveAll(tn => tn.Equals(Tokens.String));
                typeName = DetermineEvaluationTypeFromTypes(modifiedNames);
            }

            return typeName;
        }

        private string DetermineEvaluationTypeFromTypes( IEnumerable<string> typeList)
        {
            if (!typeList.Any())
            {
                return string.Empty;
            }
            //To select "String" or "Currency", all types in the typelist must match
            if (typeList.All(tn => tn.Equals(typeList.First())))
            {
                return typeList.First();
            }

            var nextType = new string[] { Tokens.Long, Tokens.LongLong, Tokens.Integer, Tokens.Byte, Tokens.Boolean };
            var result = typeList.All(tn => nextType.Contains(tn));
            if (result)
            {
                return Tokens.Long;
            }

            nextType = new string[] { Tokens.Long, Tokens.LongLong, Tokens.Integer, Tokens.Byte, Tokens.Boolean, Tokens.Single, Tokens.Double };
            result = typeList.All(tn => nextType.Contains(tn));
            if (result)
            {
                return Tokens.Double;
            }
            return string.Empty;
        }

        private bool IsLogicalContext<T>(T child)
        {
            return child is RelationalOpContext
                || child is LogicalXorOpContext
                || child is LogicalAndOpContext
                || child is LogicalOrOpContext
                || child is LogicalEqvOpContext
                || child is LogicalNotOpContext;
        }

        private bool EqualsTrueFalseLiteral<T>(T child)
        {
            if (child is VBAParser.LiteralExprContext)
            {
                var litExpr = child as VBAParser.LiteralExprContext;
                return litExpr.GetText().Equals(Tokens.True) || litExpr.GetText().Equals(Tokens.False);
            }
            return false;
        }

        #region UnreachableCaseInspectionListener
        public class UnreachableCaseInspectionListener : VBAParserBaseListener, IInspectionListener
        {
            private readonly List<QualifiedContext<ParserRuleContext>> _contexts = new List<QualifiedContext<ParserRuleContext>>();
            public IReadOnlyList<QualifiedContext<ParserRuleContext>> Contexts => _contexts;

            public QualifiedModuleName CurrentModuleName { get; set; }

            public void ClearContexts()
            {
                _contexts.Clear();
            }

            public override void EnterSelectCaseStmt([NotNull] SelectCaseStmtContext context)
            {
                _contexts.Add(new QualifiedContext<ParserRuleContext>(CurrentModuleName, context));
            }
        }
        #endregion
    }
}