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
        private List<long> _priorLongCaseClauseValues = new List<long>();
        private List<string> _priorStringCaseClauseValues = new List<string>();

        private readonly string _unreachableCaseInspectionResultFormat = "Unreachable or conflicting Case block";
        public UnreachableCaseInspection(RubberduckParserState state)
            : base(state, CodeInspectionSeverity.Suggestion) { }

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
            foreach( var selectStmt in selectCaseContexts)
            {
                inspResults.AddRange(GetUnreachableCaseBlocks(selectStmt));
            }
            return inspResults;
        }

        private IEnumerable<IInspectionResult> GetUnreachableCaseBlocks(QualifiedContext<ParserRuleContext> selectCaseStmt)
        {
            _priorLongCaseClauseValues = new List<long>();
            _priorStringCaseClauseValues = new List<string>();
            var theRef = GetTheSelectCaseReference(selectCaseStmt);
            var theDec = theRef.Declaration;
            _typeName = theDec.AsTypeName;

            var caseClauses = ParserRuleContextHelper.GetDescendents<VBAParser.CaseClauseContext>(selectCaseStmt.Context);

            var qualifiedCaseClauses = new List<QualifiedContext<ParserRuleContext>>();
            caseClauses.ForEach(clause => qualifiedCaseClauses.Add(new QualifiedContext<ParserRuleContext>(selectCaseStmt.MemberName, clause as ParserRuleContext)));

            return InspectCaseClauses(selectCaseStmt, qualifiedCaseClauses);
        }

        private IdentifierReference GetTheSelectCaseReference(QualifiedContext<ParserRuleContext> selectCaseStmt)
        {
            var selectCaseExpr = ParserRuleContextHelper.GetChild<VBAParser.SelectExpressionContext>(selectCaseStmt.Context);

            var allRefs = new List<IdentifierReference>();
            foreach (var dec in State.DeclarationFinder.MatchName(selectCaseExpr.GetText()))
            {
                allRefs.AddRange(dec.References);
            }

            var theRef = allRefs.Where(rf => ParserRuleContextHelper.HasParent(rf.Context, selectCaseStmt.Context));

            Debug.Assert(theRef.Count() == 1);

            return theRef.First();
        }

        private List<IInspectionResult> InspectCaseClauses(QualifiedContext<ParserRuleContext> selectCaseStmt, List<QualifiedContext<ParserRuleContext>> caseClauses)
        {
            Func<string, string, bool> compareEqualStrings = delegate (string s, string t)
            { return s.Equals(t, StringComparison.InvariantCulture); };

            Func<long, long, bool> compareEqualLongs = delegate (long s, long t)
            { return s == t; };

            Func<long, long, bool> isGreaterThanLong = delegate (long s, long test)
            { return s > test; };

            var stringClauseEvaluations = new List<KeyValuePair<string, Func<string, string, bool>>>();
            var longClauseEvaluations = new List<KeyValuePair<long, Func<long, long, bool>>>();

            var inspResults = new List<IInspectionResult>();
 
            foreach (var caseClause in caseClauses)
            {
                var rangeClauses = ParserRuleContextHelper.GetDescendents<VBAParser.RangeClauseContext>(caseClause.Context as ParserRuleContext);
                foreach (var ctxt in rangeClauses )
                {
         
                    if (_typeName.Equals("String"))
                    {
                        var literalExpressionContexts = ParserRuleContextHelper.GetDescendents<VBAParser.LiteralExpressionContext>(ctxt as ParserRuleContext);
                        foreach (var child in literalExpressionContexts)
                        {
                            var text = child.GetText();
                            bool addInspResult = false;
                            foreach(var cTest in stringClauseEvaluations)
                            {
                                if(cTest.Value(cTest.Key, text))
                                {
                                    addInspResult = true;
                                }
                            }
                            if (addInspResult ) //PriorCaseClauseContainsValue(text))
                            {
                                inspResults = AddInspectionResult(caseClause, inspResults);
                            }
                            else
                            {
                                var clauseTest = new KeyValuePair<string, Func<string, string, bool>>(text, compareEqualStrings);
                                stringClauseEvaluations.Add(clauseTest);
                                AddPriorCaseClauseValue(text);
                            }
                        }
                    }
                    if (_typeName.Equals("Long"))
                    {
                        bool hasStartEndContexts = false;
                        var selectStartValue = ParserRuleContextHelper.GetChild<VBAParser.SelectStartValueContext>(ctxt);
                        if (selectStartValue != null)
                        {
                            hasStartEndContexts = true;
                            bool unparsableLong = false;
                            var selectEndValue = ParserRuleContextHelper.GetChild<VBAParser.SelectEndValueContext>(ctxt);
                            long startSelectValue;
                            if (!long.TryParse(selectStartValue.GetText(), out startSelectValue))
                            {
                                unparsableLong = true;
                                inspResults = AddInspectionResult(caseClause, inspResults);
                            }
                            long endSelectValue;
                            if (!long.TryParse(selectEndValue.GetText(), out endSelectValue))
                            {
                                unparsableLong = true;
                                inspResults = AddInspectionResult(caseClause, inspResults);
                            }
                            if (!unparsableLong)
                            {

                                for (long selectVal = startSelectValue; selectVal <= endSelectValue; selectVal++)
                                {
                                    bool addInspResult = false;
                                    foreach (var cTest in longClauseEvaluations)
                                    {
                                        if (cTest.Value(cTest.Key, selectVal))
                                        {
                                            addInspResult = true;
                                        }
                                    }
                                    if (addInspResult ) //PriorCaseClauseContainsValue(selectVal))
                                    {
                                        inspResults = AddInspectionResult(caseClause, inspResults);
                                    }
                                    else
                                    {
                                        var clauseTest = new KeyValuePair<long, Func<long, long, bool>>(selectVal, compareEqualLongs);
                                        longClauseEvaluations.Add(clauseTest);
                                        AddPriorCaseClauseValue(selectVal);
                                    }
                                }
                            }
                        }

                        //bool hasComparisonOperatorContexts = false;
                        //var comparisonOperatorContexts = ParserRuleContextHelper.GetDescendents<VBAParser.ComparisonOperatorContext>(ctxt as ParserRuleContext);
                        //if (comparisonOperatorContexts.Any() && !hasStartEndContexts)
                        //{
                        //    hasComparisonOperatorContexts = true;
                        //    foreach (var child in comparisonOperatorContexts)
                        //    {
                        //        string theOp = child.GetText();
                        //        var theValue = ParserRuleContextHelper.GetDescendent<VBAParser.LExprContext>(child.Parent);
                        //        if (PriorCaseClauseContainsValue(theValue))
                        //        {
                        //            inspResults = AddInspectionResult(caseClause, inspResults);
                        //        }
                        //        else
                        //        {
                        //            AddPriorCaseClauseValue(theValue);
                        //        }
                        //    }
                        //}

                        var literalExpressionContexts = ParserRuleContextHelper.GetDescendents<VBAParser.LiteralExpressionContext>(ctxt as ParserRuleContext);
                        if( literalExpressionContexts.Any() & !hasStartEndContexts)
                        {
                            foreach (var child in literalExpressionContexts)
                            {
                                long theValue;
                                if (!long.TryParse(child.GetText(), out theValue))
                                {
                                    inspResults = AddInspectionResult(caseClause, inspResults);
                                }
                                else
                                {
                                    bool addInspResult = false;
                                    foreach (var cTest in longClauseEvaluations)
                                    {
                                        if (cTest.Value(cTest.Key, theValue))
                                        {
                                            addInspResult = true;
                                        }
                                    }
                                    if (addInspResult ) //PriorCaseClauseContainsValue(theValue))
                                    {
                                        inspResults = AddInspectionResult(caseClause, inspResults);
                                    }
                                    else
                                    {
                                        var clauseTest = new KeyValuePair<long, Func<long, long, bool>>(theValue, compareEqualLongs);
                                        longClauseEvaluations.Add(clauseTest);
                                        AddPriorCaseClauseValue(theValue);
                                    }
                                }
                            }
                        }
                    }
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

        private void AddPriorCaseClauseValue(long theValue)
        {
            _priorLongCaseClauseValues.Add(theValue);
        }

        private bool PriorCaseClauseContainsValue(long theValue)
        {
            return _priorLongCaseClauseValues.Contains(theValue);
        }

        private void AddPriorCaseClauseValue(string theValue)
        {
            _priorStringCaseClauseValues.Add(theValue);
        }

        private bool PriorCaseClauseContainsValue(string theValue)
        {
            return _priorStringCaseClauseValues.Contains(theValue);
        }

        private void VBATypes()
        {
            string[] types = {"Boolean", "Integer",
                    "Long", "Single","Double(negative)",
                    "Double(positive)", "Currency",
            "Date", "String", "Object", "Variant", "User defined" };
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
