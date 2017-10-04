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
        private static long BYTEMAX => 255;
        private static long BYTEMIN => 0;

        private static long INTMAX => 32767;
        private static long INTMIN => -32768;

        private static long LONGMAX => 2147486647;
        private static long LONGMIN => -2147486648;

        private static decimal CURRENCYMAX => 922337203685477.5807M;
        private static decimal CURRENCYMIN => -922337203685477.5808M;

        private VBAParser.CaseElseClauseContext _caseElseClause;
        QualifiedModuleName _qMemberName;
        //private bool _hasCaseElseClause;

        private readonly string _unreachableCaseInspectionResultFormat = "Unreachable or conflicting Case block";
        public UnreachableCaseInspection(RubberduckParserState state)
            : base(state, CodeInspectionSeverity.Suggestion)
        {
            //_hasCaseElseClause = false;
        }

        public override IInspectionListener Listener { get; } =
            new UnreachableCaseInspectionListener();

        public override Type Type => typeof(UnreachableCaseInspection);

        public override CodeInspectionType InspectionType => CodeInspectionType.CodeQualityIssues;


        public override IEnumerable<IInspectionResult> GetInspectionResults()
        {
            var selectCaseContexts = Listener.Contexts
                .Where(result => !IsIgnoringInspectionResultFor(result.ModuleName, result.Context.Start.Line));
            //.Select(result => new QualifiedContextInspectionResult(this,
            //                                        _unreachableCaseInspectionResultFormat,
            //                                        result));


            var inspResults = new List<IInspectionResult>();
            foreach (var selectStmt in selectCaseContexts)
            {
                inspResults.AddRange(GetUnreachableCaseBlocks(selectStmt));
            }
            return inspResults;
        }

        private IEnumerable<IInspectionResult> GetUnreachableCaseBlocks(QualifiedContext<ParserRuleContext> selectCaseStmt)
        {
            var theRef = GetTheSelectCaseReference(selectCaseStmt);

            //For now we only handle SelectCaseStmt that use a simple variable reference
            if(theRef == null)
            {
                return new List<IInspectionResult>();
            }

            var caseClauses = ParserRuleContextHelper.GetChildren<VBAParser.CaseClauseContext>(selectCaseStmt.Context);

            _caseElseClause = ParserRuleContextHelper.GetChild<VBAParser.CaseElseClauseContext>(selectCaseStmt.Context);
            //_hasCaseElseClause = _caseElseClause != null;

            var qualifiedCaseClauses = new List<QualifiedContext<ParserRuleContext>>();
            _qMemberName = selectCaseStmt.ModuleName;
            caseClauses.ForEach(clause => qualifiedCaseClauses.Add(new QualifiedContext<ParserRuleContext>(selectCaseStmt.ModuleName, clause as ParserRuleContext)));

            return HandleSelectCase(qualifiedCaseClauses, theRef);
        }

        private IdentifierReference GetTheSelectCaseReference(QualifiedContext<ParserRuleContext> selectCaseStmt)
        {
            var selectCaseExpr = ParserRuleContextHelper.GetChild<VBAParser.SelectExpressionContext>(selectCaseStmt.Context);

            var allRefs = new List<IdentifierReference>();
            var name = selectCaseExpr.GetText();
            var test = State.DeclarationFinder.MatchName(selectCaseExpr.GetText());
            foreach (var dec in State.DeclarationFinder.MatchName(selectCaseExpr.GetText()))
            {
                allRefs.AddRange(dec.References);
            }

            if (!allRefs.Any())
            {
                return null;
            }

            var selectCaseReference = allRefs.Where(rf => ParserRuleContextHelper.HasParent(rf.Context, selectCaseStmt.Context)
                                    && (ParserRuleContextHelper.HasParent(rf.Context, selectCaseExpr)));

            //Debug.Assert(selectCaseReference.Count() == 1);
            if (selectCaseReference.Count() != 1)
            {
                int i = 0;
            }
            return selectCaseReference.First();
        }

        //TODO: Check for matching types early on
        //TODO: Check that all contexts are parseable
        private List<IInspectionResult> HandleSelectCase(List<QualifiedContext<ParserRuleContext>> caseClauses, IdentifierReference theRef)
        {
            var inspResults = new List<IInspectionResult>();

            var rangeEvals = LoadBoundaryEvaluations(theRef.Declaration.AsTypeName, new List<IComparable>());

            foreach (var caseClause in caseClauses)
            {
                var rangeClauses = ParserRuleContextHelper.GetChildren<VBAParser.RangeClauseContext>(caseClause.Context as ParserRuleContext);
                foreach (var ctxt in rangeClauses)
                {
                    var test = ctxt.GetText();
                    var rangeClause = new RangeClause(ctxt, theRef);
                    if (!rangeClause.IsParseable)
                    {
                        inspResults = AddInspectionResult(caseClause, inspResults);
                        continue;
                    }

                    if (rangeEvals.Any())
                    {
                        bool hasConflict = false;
                        for (int idx = 0; idx < rangeEvals.Count() && !hasConflict; idx++)
                        {
                            hasConflict = rangeClause.CompareTo(rangeEvals[idx]) == 0;
                        }

                        if (hasConflict)
                        {
                            inspResults = AddInspectionResult(caseClause, inspResults);
                        }
                    }
                    rangeEvals.Add(rangeClause);
                }
            }

            if (_caseElseClause != null && theRef.Declaration.AsTypeName.Equals("Boolean"))
            {
                //Check if at least one Case exists for both True and False, then the Case Else clause is unreachable
                bool hasTrueResult = false;
                bool hasFalseResult = false;
                for (int idx = 0; idx < rangeEvals.Count() && !(hasTrueResult && hasFalseResult); idx++)
                {
                    var rangeEval = rangeEvals[idx];
                    var resultTrue = rangeEval.CompareTo(new RangeClauseExtent<decimal>(1, theRef.Declaration.AsTypeName, "=")) == 0;
                    if (resultTrue)
                    {
                        hasTrueResult = true;
                    }
                    var resultFalse = rangeEval.CompareTo(new RangeClauseExtent<decimal>(0, theRef.Declaration.AsTypeName, "=")) == 0;
                    if (resultFalse)
                    {
                        hasFalseResult = true;
                    }
                }
                if (hasTrueResult && hasFalseResult)
                {
                    inspResults = AddInspectionResult(new QualifiedContext<ParserRuleContext>(_qMemberName, _caseElseClause as ParserRuleContext), inspResults);
                }
            }

            return inspResults;
        }

        private List<IInspectionResult> AddInspectionResult(QualifiedContext<ParserRuleContext> result, List<IInspectionResult> inspResults)
        {
            inspResults.Add(new QualifiedContextInspectionResult(this,
                    _unreachableCaseInspectionResultFormat,
                    result));
            return inspResults;
        }

        private static List<IComparable> LoadBoundaryEvaluations(string theTypeName, List<IComparable> rangeEvals)
        {
            if (theTypeName.Equals("Integer"))
            {
                rangeEvals.Add(new RangeClauseExtent<long>(INTMAX, theTypeName, ">"));
                rangeEvals.Add(new RangeClauseExtent<long>(INTMIN, theTypeName, "<"));
            }
            else if (theTypeName.Equals("Byte"))
            {
                rangeEvals.Add(new RangeClauseExtent<long>(BYTEMAX, theTypeName, ">"));
                rangeEvals.Add(new RangeClauseExtent<long>(BYTEMIN, theTypeName, "<"));
            }
            else if (theTypeName.Equals("Currency"))
            {
                rangeEvals.Add(new RangeClauseExtent<decimal>(CURRENCYMAX, theTypeName, ">"));
                rangeEvals.Add(new RangeClauseExtent<decimal>(CURRENCYMIN, theTypeName, "<"));
            }
            else if (theTypeName.Equals("Boolean"))
            {
                rangeEvals.Add(new RangeClauseExtent<decimal>(1, theTypeName, ">"));
                rangeEvals.Add(new RangeClauseExtent<decimal>(0, theTypeName, "<"));
            }

            //TODO: Add Single, Decimal
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
