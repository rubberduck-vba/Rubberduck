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

namespace Rubberduck.Inspections.Concrete
{
    public sealed class UnreachableCaseInspection : ParseTreeInspectionBase
    {
        public enum ClauseEvaluationResult { Unreachable, MismatchType, CaseElse };

        internal Dictionary<ClauseEvaluationResult, string> ResultMessages = new Dictionary<ClauseEvaluationResult, string>()
        {
            [ClauseEvaluationResult.Unreachable] = InspectionsUI.UnreachableCaseInspection_Unreachable,
            [ClauseEvaluationResult.MismatchType] = InspectionsUI.UnreachableCaseInspection_TypeMismatch,
            [ClauseEvaluationResult.CaseElse] = InspectionsUI.UnreachableCaseInspection_CaseElse
        };

        public UnreachableCaseInspection(RubberduckParserState state) : base(state) { }

        public override IInspectionListener Listener { get; } =
            new UnreachableCaseInspectionListener();

        protected override IEnumerable<IInspectionResult> DoGetInspectionResults()
        {
            var inspResults = new List<IInspectionResult>();

            var selectCaseStmtContexts = Listener.Contexts
                .Where(result => !IsIgnoringInspectionResultFor(result.ModuleName, result.Context.Start.Line));

            foreach(var selectCaseStmt in selectCaseStmtContexts)
            {
                var selectStmtWrapper = new SelectStatementInspectionWrapper(selectCaseStmt, State);
                if (selectStmtWrapper.CanBeInspected)
                {
                    inspResults.AddRange(InspectSelectStatement(selectStmtWrapper));
                }
            }
            return inspResults;
        }

        private IEnumerable<IInspectionResult> InspectSelectStatement(SelectStatementInspectionWrapper selectStmt)
        {
            var contextValues = new ContextValueVisitor(State, selectStmt.EvaluationTypeName);
            selectStmt.QualifiedContext.Context.Accept(contextValues);
            return InspectSelectStatementAsType(selectStmt, contextValues);
        }

        private IEnumerable<IInspectionResult> InspectSelectStatementAsType(SelectStatementInspectionWrapper selectStmt, ContextValueVisitor contextValues)
        {
            var contextValueResults_long = contextValues.ResultsAsLong();
            if (contextValues.EvaluationTypeName.Equals(Tokens.Long))
            {
                contextValueResults_long.SetExtents(CompareExtents.LONGMIN, CompareExtents.LONGMAX);
                return InspectSelectStatement(selectStmt, contextValueResults_long);
            }
            else if (contextValues.EvaluationTypeName.Equals(Tokens.Int) || contextValues.EvaluationTypeName.Equals(Tokens.Integer))
            {
                contextValueResults_long.SetExtents(CompareExtents.INTEGERMIN, CompareExtents.INTEGERMAX);
                return InspectSelectStatement(selectStmt, contextValueResults_long);
            }
            else if (contextValues.EvaluationTypeName.Equals(Tokens.Double))
            {
                var typeSpecific = contextValues.ResultsAsDouble();
                return InspectSelectStatement(selectStmt, typeSpecific);
            }
            else if (contextValues.EvaluationTypeName.Equals(Tokens.Single))
            {
                var typeSpecific = contextValues.ResultsAsDouble();
                typeSpecific.SetExtents(CompareExtents.SINGLEMIN, CompareExtents.SINGLEMAX);
                return InspectSelectStatement(selectStmt, typeSpecific);
            }
            else if (contextValues.EvaluationTypeName.Equals(Tokens.Currency))
            {
                var typeSpecific = contextValues.ResultsAsCurrency();
                typeSpecific.SetExtents(CompareExtents.CURRENCYMIN, CompareExtents.CURRENCYMAX);
                return InspectSelectStatement(selectStmt, typeSpecific);
            }
            else if (contextValues.EvaluationTypeName.Equals(Tokens.Byte))
            {
                contextValueResults_long.SetExtents(CompareExtents.BYTEMIN, CompareExtents.BYTEMAX);
                return InspectSelectStatement(selectStmt, contextValueResults_long);
            }
            else if (contextValues.EvaluationTypeName.Equals(Tokens.Boolean))
            {
                var typeSpecific = contextValues.ResultsAsBoolean();
                return InspectSelectStatement(selectStmt, typeSpecific);
            }
            else if (contextValues.EvaluationTypeName.Equals(Tokens.String))
            {
                var typeSpecific = contextValues.ResultsAsString();
                return InspectSelectStatement(selectStmt, typeSpecific);
            }
            return new List<IInspectionResult>();
        }

        private IEnumerable<IInspectionResult>
        InspectSelectStatement<T>(SelectStatementInspectionWrapper selectStmt, ContextValueResults<T> ctxtValueResults) where T : IComparable<T>
        {
            var inspResults = new List<IInspectionResult>();
            var cummulativeCoverage = new SummaryCoverage<T>(ctxtValueResults.Extents);

            foreach (var caseClause in selectStmt.CaseClauses)
            {
                if (cummulativeCoverage.CoversAllValues)
                {
                    var result = CreateInspectionResult(selectStmt.QualifiedContext, caseClause, ResultMessages[ClauseEvaluationResult.Unreachable]);
                    inspResults.Add(result);
                    continue;
                }

                var caseClauseWrapper = new CaseClauseWrapper<T>(caseClause, ctxtValueResults);
                if (caseClauseWrapper.CanBeInspected)
                {
                    var additionalCoverage = caseClauseWrapper.RemoveCoverageRedundantTo(cummulativeCoverage);
                    if (additionalCoverage.Empty)
                    {
                        var result = CreateInspectionResult(selectStmt.QualifiedContext, caseClause, ResultMessages[ClauseEvaluationResult.Unreachable]);
                        inspResults.Add(result);
                    }
                    else
                    {
                        cummulativeCoverage.Add(additionalCoverage);
                    }
                }
                else
                {
                    if (caseClauseWrapper.IsMismatch)
                    {
                        var result = CreateInspectionResult(selectStmt.QualifiedContext, caseClause, ResultMessages[ClauseEvaluationResult.MismatchType]);
                        inspResults.Add(result);
                    }
                }
            }

            if (cummulativeCoverage.CoversAllValues && selectStmt.HasCaseElse)
            {
                var result = CreateInspectionResult(selectStmt.QualifiedContext, selectStmt.CaseElse, ResultMessages[ClauseEvaluationResult.CaseElse]);
                inspResults.Add(result);
            }
            return inspResults;
        }

        private IInspectionResult CreateInspectionResult(QualifiedContext<ParserRuleContext> selectStmt, ParserRuleContext unreachableBlock, string message)
        {
            return new QualifiedContextInspectionResult(this,
                        message,
                        new QualifiedContext<ParserRuleContext>(selectStmt.ModuleName, unreachableBlock));
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

            public override void EnterSelectCaseStmt([NotNull] VBAParser.SelectCaseStmtContext context)
            {
                _contexts.Add(new QualifiedContext<ParserRuleContext>(CurrentModuleName, context));
            }
        }
        #endregion
    }
}