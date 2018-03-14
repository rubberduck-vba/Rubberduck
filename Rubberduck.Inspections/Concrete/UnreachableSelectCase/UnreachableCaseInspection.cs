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

namespace Rubberduck.Inspections.Concrete.UnreachableSelectCase
{
    public sealed class UnreachableCaseInspection : ParseTreeInspectionBase
    {
        private readonly ISummaryCoverageFactory _summaryCoverageFactory;
        private readonly IUnreachableCaseInspectionVisitorFactory _visitorFactory;
        private enum ClauseEvaluationResult { Unreachable, MismatchType, CaseElse };

        private Dictionary<ClauseEvaluationResult, string> ResultMessages = new Dictionary<ClauseEvaluationResult, string>()
        {
            [ClauseEvaluationResult.Unreachable] = InspectionsUI.UnreachableCaseInspection_Unreachable,
            [ClauseEvaluationResult.MismatchType] = InspectionsUI.UnreachableCaseInspection_TypeMismatch,
            [ClauseEvaluationResult.CaseElse] = InspectionsUI.UnreachableCaseInspection_CaseElse
        };

        public UnreachableCaseInspection(RubberduckParserState state/*, ISummaryCoverageFactory factory, IUnreachableCaseInspectionVisitorFactory visitorFactory*/) : base(state)
        {
            //TODO_Question: Candidates for IoCInstaller?  Or...not appropriate?
            _summaryCoverageFactory = new SummaryCoverageFactory();
            _visitorFactory = new UnreachableCaseInspectionVisitorFactory();
        }

        public override IInspectionListener Listener { get; } =
            new UnreachableCaseInspectionListener();

        public UnreachableCaseInspectionMismatchListener MismatchListener { get; } =
            new UnreachableCaseInspectionMismatchListener();

        private List<IInspectionResult> InspectionResults { set; get; } = new List<IInspectionResult>();

        protected override IEnumerable<IInspectionResult> DoGetInspectionResults()
        {
            var selectCaseStmtContexts = Listener.Contexts
                .Where(result => !IsIgnoringInspectionResultFor(result.ModuleName, result.Context.Start.Line));

            InspectionResults = new List<IInspectionResult>();
            var selectCasesToInspect = new List<IUnreachableCaseInspectionSelectStmt>();
            foreach (var ctxt in selectCaseStmtContexts)
            {
                IParseTreeVisitor<IUnreachableCaseInspectionValue> ptVisitor = new UnreachableCaseInspectionValueVisitor(State, new IUnreachableCaseInspectionValueFactory());
                var inspectionSelectCaseStmt = new UnreachableCaseInspectionSelectStmtContext(ctxt, ptVisitor);
                if (inspectionSelectCaseStmt.CanBeInspected)
                {
                    InspectSelectStatement(inspectionSelectCaseStmt);
                }
            }
            return InspectionResults;
        }


        private void InspectSelectStatement(IUnreachableCaseInspectionSelectStmt selectStmt)
        {
            var qualifiedSelectCaseContext = selectStmt.QualifiedContext;
            var cummulativeCoverage = _summaryCoverageFactory.Create(selectStmt.EvaluationTypeName);

            var selectCaseStatementContext = (VBAParser.SelectCaseStmtContext)qualifiedSelectCaseContext.Context;
            foreach (var inspCaseClause in selectStmt.CaseClauses)
            {
                if (cummulativeCoverage.CoversAllValues)
                {
                    //Once all values are covered, the remaining CaseClauses are now unreachable
                    CreateInspectionResult(qualifiedSelectCaseContext, inspCaseClause.Context, ResultMessages[ClauseEvaluationResult.Unreachable]);
                    continue;
                }

                //MismatchListener.ClearContexts();
                //var caseClauseVisitor = PrepareCaseClauseSummaryVisitor(inspCaseClause, selectStmt.EvaluationTypeName);
                var caseClauseCoverge = RetrieveSummaryCoverage(inspCaseClause, selectStmt.EvaluationTypeName);
                //RetrieveSummaryCoverage
                //var caseClauseCoverge = inspCaseClause.Context.Accept(caseClauseVisitor);
                //var isMismatch = MismatchListener.Contexts.Count == inspCaseClause.Context.GetDescendents<VBAParser.RangeClauseContext>().Count();
                if(!MismatchListener.MismatchFound(inspCaseClause))
                {
                    if (caseClauseCoverge.HasClausesNotCoveredBy(cummulativeCoverage, out ISummaryCoverage additionalCoverage))
                    {
                        cummulativeCoverage.Add(additionalCoverage);
                    }
                    else
                    {
                        //If there are no conditions to add, then the current CaseClause's conditions are covered
                        //by the combination of prior CaseClauses - and the CaseClause is therefore unreachable
                        CreateInspectionResult(qualifiedSelectCaseContext, inspCaseClause.Context, ResultMessages[ClauseEvaluationResult.Unreachable]);
                    }
                }
                else
                {
                    //Call out CaseClauses that cannot be implicitly converted to the SelectCase type as a special case of unreachable
                    CreateInspectionResult(qualifiedSelectCaseContext, inspCaseClause.Context, ResultMessages[ClauseEvaluationResult.MismatchType]);
                }
            }

            if (cummulativeCoverage.CoversAllValues && !(selectCaseStatementContext.caseElseClause() is null))
            {
                CreateInspectionResult(qualifiedSelectCaseContext, selectCaseStatementContext.caseElseClause(), ResultMessages[ClauseEvaluationResult.CaseElse]);
            }
        }

        private ISummaryCoverage RetrieveSummaryCoverage(IUnreachableCaseInspectionCaseClause inspCaseClause, string evalTypeName)
        {
            MismatchListener.ClearContexts();
            var caseClauseVisitor = new CaseClauseSummaryVisitor((VBAParser.CaseClauseContext)inspCaseClause.Context, State, evalTypeName);
            caseClauseVisitor.IncompatibleRangeClauseDetected += MismatchListener.IncompatibleRangeDetected;
            var caseClauseCoverge = inspCaseClause.Context.Accept(caseClauseVisitor);
            return caseClauseCoverge;
        }

        private void CreateInspectionResult(QualifiedContext<ParserRuleContext> selectStmt, ParserRuleContext unreachableBlock, string message)
        {
            var result = new QualifiedContextInspectionResult(this,
                        message,
                        new QualifiedContext<ParserRuleContext>(selectStmt.ModuleName, unreachableBlock));
            InspectionResults.Add(result);
        }

        #region UnreachableCaseInspectionListeners
        public class UnreachableCaseInspectionMismatchListener
        {
            private readonly List<ParserRuleContext> _contexts = new List<ParserRuleContext>();
            public IReadOnlyList<ParserRuleContext> Contexts => _contexts;

            public bool MismatchFound(IUnreachableCaseInspectionCaseClause inspCaseClause)
            {
                return Contexts.Count == inspCaseClause.Context.GetDescendents<VBAParser.RangeClauseContext>().Count();
            }

            public QualifiedModuleName CurrentModuleName { get; set; }

            public void ClearContexts()
            {
                _contexts.Clear();
            }

            public void IncompatibleRangeDetected(object sender, IncompatibleRangeClauseDetectedArgs e)
            {
                _contexts.Add(e.RangeClause);
            }
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
        #endregion
    }
}