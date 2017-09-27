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
        private string _typeName;

        private static long BYTEMAX => 255;
        private static long BYTEMIN => 0;

        private static long INTMAX => 32767;
        private static long INTMIN => -32768;

        private static long LONGMAX => 2147486647;
        private static long LONGMIN => -2147486648;

        private readonly string _unreachableCaseInspectionResultFormat = "Unreachable or conflicting Case block";
        public UnreachableCaseInspection(RubberduckParserState state)
            : base(state, CodeInspectionSeverity.Suggestion)
        {
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

            if(theRef == null)
            {
                return new List<IInspectionResult>();
            }

            var caseClauses = ParserRuleContextHelper.GetChildren<VBAParser.CaseClauseContext>(selectCaseStmt.Context);

            var qualifiedCaseClauses = new List<QualifiedContext<ParserRuleContext>>();
            caseClauses.ForEach(clause => qualifiedCaseClauses.Add(new QualifiedContext<ParserRuleContext>(selectCaseStmt.ModuleName, clause as ParserRuleContext)));

            return HandleSelectCase(qualifiedCaseClauses, theRef);
        }

        //private bool IsConstantExpression(VBAParser.SelectExpressionContext selectCaseExpr)
        //{
        //    return selectCaseExpr.ChildCount == 1 && selectCaseExpr.children[0] is VBAParser.LiteralExprContext;
        //}

        private IdentifierReference GetTheSelectCaseReference(QualifiedContext<ParserRuleContext> selectCaseStmt)
        {
            var selectCaseExpr = ParserRuleContextHelper.GetChild<VBAParser.SelectExpressionContext>(selectCaseStmt.Context);
            //if (IsConstantExpression(selectCaseExpr))
            //{
            //    return null;
            //}

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

        private List<IInspectionResult> HandleSelectCase(List<QualifiedContext<ParserRuleContext>> caseClauses, IdentifierReference theRef)
        {
            var inspResults = new List<IInspectionResult>();
            var rangeEvals = new List<IComparable>();
            _typeName = theRef.Declaration.AsTypeName;

            if (_typeName.Equals("Integer"))
            {
                rangeEvals.Add(new RangeClauseExtent<long>(INTMAX, "Integer", ">"));
                rangeEvals.Add(new RangeClauseExtent<long>(INTMIN, "Integer", "<"));
            }

            if (_typeName.Equals("Byte"))
            {
                rangeEvals.Add(new RangeClauseExtent<long>(BYTEMAX, "Byte", ">"));
                rangeEvals.Add(new RangeClauseExtent<long>(BYTEMIN, "Byte", "<"));
            }

            foreach (var caseClause in caseClauses)
            {
                var rangeClauses = ParserRuleContextHelper.GetChildren<VBAParser.RangeClauseContext>(caseClause.Context as ParserRuleContext);
                foreach (var ctxt in rangeClauses)
                {
                    var test = ctxt.GetText();
                    //TODO: Check for ctxt type to match _typeName here
                    //TODO: Check that ctxt is parse-able to the _typeName type
                    var rangeClause = new RangeClause(ctxt, theRef);
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

            return inspResults;
        }

        private List<IInspectionResult> AddInspectionResult(QualifiedContext<ParserRuleContext> result, List<IInspectionResult> inspResults)
        {
            inspResults.Add(new QualifiedContextInspectionResult(this,
                    _unreachableCaseInspectionResultFormat,
                    result));
            return inspResults;
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
