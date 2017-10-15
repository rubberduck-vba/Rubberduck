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
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;

namespace Rubberduck.Inspections.Concrete
{
    internal class UnreachableCaseInspection : ParseTreeInspectionBase
    {
        //TODO: Add replace with UI Resource
        private readonly string _unreachableCaseInspectionResultFormat = "Unreachable or conflicting Case block";

        public UnreachableCaseInspection(RubberduckParserState state)
            : base(state, CodeInspectionSeverity.Suggestion){ }

        public override IInspectionListener Listener { get; } =
            new UnreachableCaseInspectionListener();

        public override CodeInspectionType InspectionType => CodeInspectionType.CodeQualityIssues;

        protected override IEnumerable<IInspectionResult> DoGetInspectionResults()
        {
            var selectCaseContexts = Listener.Contexts
                .Where(result => !IsIgnoringInspectionResultFor(result.ModuleName, result.Context.Start.Line));

            var inspResults = new List<IInspectionResult>();
            foreach (var selectStmt in selectCaseContexts)
            {
                var unreachableCaseBlocks = GetUnreachableCaseBlocks(selectStmt);

                unreachableCaseBlocks.ForEach(unreachableBlock => inspResults.Add(CreateInspectionResult(selectStmt, unreachableBlock)));
            }
            return inspResults;
        }

        private IInspectionResult CreateInspectionResult(QualifiedContext<ParserRuleContext> selectStmt, ParserRuleContext unreachableBlock)
        {
            return new QualifiedContextInspectionResult(this,
                        _unreachableCaseInspectionResultFormat,
                        new QualifiedContext<ParserRuleContext>(selectStmt.ModuleName, unreachableBlock));
        }

        private List<ParserRuleContext> GetUnreachableCaseBlocks(QualifiedContext<ParserRuleContext> selectCaseStmt)
        {
            var selectCaseExpr = ParserRuleContextHelper.GetChild<VBAParser.SelectExpressionContext>(selectCaseStmt.Context);
            if(selectCaseExpr == null)
            {
                return new List<ParserRuleContext>();
            }

            IdentifierReference selectCaseIdentifierReference = null;
            var typeName = string.Empty;

            if (IsBooleanSelectExpression(selectCaseExpr))
            {
                typeName = "Boolean";
                return EvaluateCaseClauses(selectCaseStmt, selectCaseIdentifierReference, typeName);
            }

            var smplName = ParserRuleContextHelper.GetDescendent<VBAParser.SimpleNameExprContext>(selectCaseExpr);
            if (smplName != null)
            {
                selectCaseIdentifierReference = GetTheSelectCaseReference(selectCaseStmt, smplName.GetText());
                if (selectCaseIdentifierReference != null)
                {
                    typeName = selectCaseIdentifierReference.Declaration.AsTypeName;
                    return EvaluateCaseClauses(selectCaseStmt, selectCaseIdentifierReference, typeName);
                }
            }

            return new List<ParserRuleContext>();
        }

        private bool IsBooleanSelectExpression(VBAParser.SelectExpressionContext selectExpr)
        {
            var relationalOpCtxt = ParserRuleContextHelper.GetChild<VBAParser.RelationalOpContext>(selectExpr);
            return relationalOpCtxt != null;
        }

        private IdentifierReference GetTheSelectCaseReference(QualifiedContext<ParserRuleContext> selectCaseStmt, string theName)
        {
            var selectCaseExpr = ParserRuleContextHelper.GetChild<VBAParser.SelectExpressionContext>(selectCaseStmt.Context);

            var allRefs = new List<IdentifierReference>();
            foreach (var dec in State.DeclarationFinder.MatchName(theName))
            {
                allRefs.AddRange(dec.References);
            }

            if (!allRefs.Any())
            {
                return null;
            }

            var selectCaseReference = allRefs.Where(rf => ParserRuleContextHelper.HasParent(rf.Context, selectCaseStmt.Context)
                                    && (ParserRuleContextHelper.HasParent(rf.Context, selectCaseExpr)));

            Debug.Assert(selectCaseReference.Count() == 1);
            return selectCaseReference.First();
        }

        private List<ParserRuleContext> EvaluateCaseClauses(QualifiedContext<ParserRuleContext> selectCaseStmt, IdentifierReference theRef, string typeName)
        {
            var caseClauses = ParserRuleContextHelper.GetChildren<VBAParser.CaseClauseContext>(selectCaseStmt.Context);
            var caseElseClause = ParserRuleContextHelper.GetChild<VBAParser.CaseElseClauseContext>(selectCaseStmt.Context);

            var unreachableClauses = new List<ParserRuleContext>();

            var rangeEvals = LoadBoundaryEvaluations(typeName, new List<IRangeClause>());

            var caseElseUnreachable = false;
            foreach (var caseClause in caseClauses)
            {
                var rangeClauseContexts = ParserRuleContextHelper.GetChildren<VBAParser.RangeClauseContext>(caseClause);
                foreach (var rangeClauseCtxt in rangeClauseContexts)
                {
                    var test = rangeClauseCtxt.GetText();
                    var rangeClause = new RangeClause(State, rangeClauseCtxt, theRef, typeName);
                    if (!rangeClause.IsParseable)
                    {
                        if (!rangeClause.MatchesSelectCaseType)
                        {
                            unreachableClauses.Add(caseClause);
                        }
                        continue;
                    }

                    if (rangeEvals.Any())
                    {
                        var isReachable = true;
                        for (int idx = 0; idx < rangeEvals.Count() && isReachable; idx++)
                        {
                            isReachable = rangeClause.IsReachable(rangeEvals[idx]);
                            if (rangeClause.HasUnreachableCaseElse)
                            {
                                caseElseUnreachable = true;
                            }
                        }

                        if (!isReachable)
                        {
                            unreachableClauses.Add(caseClause);
                        }
                    }
                    rangeEvals.Add(rangeClause);
                }
            }

            if (caseElseClause != null && caseElseUnreachable)
            {
                unreachableClauses.Add(caseElseClause);
            }

            return unreachableClauses;
        }

        private List<IInspectionResult> AddInspectionResult(QualifiedContext<ParserRuleContext> result, List<IInspectionResult> inspResults)
        {
            inspResults.Add(new QualifiedContextInspectionResult(this,
                    _unreachableCaseInspectionResultFormat,
                    result));
            return inspResults;
        }

        private static List<IRangeClause> LoadBoundaryEvaluations(string theTypeName, List<IRangeClause> rangeEvals)
        {

            if (theTypeName.Equals("Long"))
            {
                long LONGMAX = 2147486647;
                long LONGMIN = -2147486648;
                return LoadExents(LONGMIN, LONGMAX, rangeEvals);
            }
            if (theTypeName.Equals("Integer"))
            {
                long INTEGERMAX = 32767;
                long INTEGERMIN = -32768;
                return LoadExents(INTEGERMIN, INTEGERMAX, rangeEvals);
            }
            else if (theTypeName.Equals("Byte"))
            {
                long BYTEMAX = 255;
                long BYTEMIN = 0;
                return LoadExents(BYTEMIN, BYTEMAX, rangeEvals);
            }
            else if (theTypeName.Equals("Boolean"))
            {
                int BOOLEANMAX = 1;
                int BOOLEANMIN = 0;
                return LoadExents(BOOLEANMIN, BOOLEANMAX, rangeEvals);
            }
            else if (theTypeName.Equals("Currency"))
            {
                decimal CURRENCYMAX = 922337203685477.5807M;
                decimal CURRENCYMIN = -922337203685477.5808M;
                return LoadExents(CURRENCYMIN, CURRENCYMAX, rangeEvals);
            }
            else if (theTypeName.Equals("Single"))
            {
                double SINGLEMAX = 3402823E38;
                double SINGLEMIN = -3402823E38;
                return LoadExents(SINGLEMIN, SINGLEMAX, rangeEvals);
            }
            //Decimal/Variant data type not supported by extent checks

            return rangeEvals;
        }

        private static List<IRangeClause> LoadExents<T>(T min, T max, List<IRangeClause> rangeEvals) where T : System.IComparable
        {
            rangeEvals.Add(new RangeClauseExtent<T>(min, "<"));
            rangeEvals.Add(new RangeClauseExtent<T>(max, ">"));
            return rangeEvals;
        }

        public class UnreachableCaseInspectionListener : VBAParserBaseListener, IInspectionListener
        {
            private readonly List<QualifiedContext<ParserRuleContext>> _contexts = new List<QualifiedContext<ParserRuleContext>>();
            public IReadOnlyList<QualifiedContext<ParserRuleContext>> Contexts => _contexts;

            public QualifiedModuleName CurrentModuleName { get; set; }

            public void ClearContexts()
            {
                _contexts.Clear();
            }

            public override void EnterSelectCaseStmt([NotNull] VBAParser.SelectCaseStmtContext context)
            {
                _contexts.Add(new QualifiedContext<ParserRuleContext>(CurrentModuleName, context));
            }
        }
    }
}
