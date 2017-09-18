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
        private Dictionary<string, Func<long, long, long, long, bool>> _operatorsLong;
        private Dictionary<string, Func<CompareParams<double>, bool>> _operatorsDouble;

        private struct CompareParams<T>
        {
            public CompareParams(T cMin, T cMax, T vMin, T vMax)
            {
                candidateMin = cMin;
                candidateMax = cMax;
                minVal = vMin;
                maxVal = vMax;
            }

            public CompareParams(T cVal, T vVal)
            {
                candidateMin = cVal;
                candidateMax = cVal;
                minVal = vVal;
                maxVal = vVal;
            }

            public T candidateMin;
            public T candidateMax;
            public T minVal;
            public T maxVal;
        }

        private readonly string _unreachableCaseInspectionResultFormat = "Unreachable or conflicting Case block";
        public UnreachableCaseInspection(RubberduckParserState state)
            : base(state, CodeInspectionSeverity.Suggestion)
        {
            _operatorsDouble = new Dictionary<string, Func<CompareParams<double>, bool>>();
            _operatorsLong = new Dictionary<string, Func<long, long, long, long, bool>>();
            _operatorsLong.Add("=", _compareCaseEqualLongs);
            _operatorsLong.Add("<>", _compareCaseNEQLongs);
            _operatorsLong.Add(">", _compareCaseGTLongs);
            _operatorsLong.Add("<", _compareCaseLTLongs);
            _operatorsLong.Add(">=", _compareCaseGTELongs);
            _operatorsLong.Add("<=", _compareCaseLTELongs);

            //_operatorsDouble = new Dictionary<string, Func<double, double, double, double, bool>>();
            //_operatorsDouble.Add("=", _compareCaseEqualDouble);
            //_operatorsDouble.Add("<>", _compareCaseNEQDouble);
            //_operatorsDouble.Add(">", _compareCaseGTDouble);
            //_operatorsDouble.Add("<", _compareCaseLTDouble);
            //_operatorsDouble.Add(">=", _compareCaseGTEDouble);
            //_operatorsDouble.Add("<=", _compareCaseLTEDouble);
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
            foreach( var selectStmt in selectCaseContexts)
            {
                inspResults.AddRange(GetUnreachableCaseBlocks(selectStmt));
            }
            return inspResults;
        }

        private IEnumerable<IInspectionResult> GetUnreachableCaseBlocks(QualifiedContext<ParserRuleContext> selectCaseStmt)
        {
            var theRef = GetTheSelectCaseReference(selectCaseStmt);
            _typeName = theRef.Declaration.AsTypeName;

            var caseClauses = ParserRuleContextHelper.GetChildren<VBAParser.CaseClauseContext>(selectCaseStmt.Context);

            var qualifiedCaseClauses = new List<QualifiedContext<ParserRuleContext>>();
            caseClauses.ForEach(clause => qualifiedCaseClauses.Add(new QualifiedContext<ParserRuleContext>(selectCaseStmt.ModuleName, clause as ParserRuleContext)));

            return InspectCaseClauses(selectCaseStmt, qualifiedCaseClauses);
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

            var selectCaseReference = allRefs.Where(rf => ParserRuleContextHelper.HasParent(rf.Context, selectCaseStmt.Context)
                                    && (ParserRuleContextHelper.HasParent(rf.Context, selectCaseExpr)));

            //Debug.Assert(selectCaseReference.Count() == 1);
            if(selectCaseReference.Count() != 1)
            {
                int i = 0;
            }
            return selectCaseReference.First();
        }

        private void VBATypes()
        {
            string[] types = {"Boolean", "Integer",
                    "Long", "Single","Double(negative)",
                    "Double(positive)", "Currency",
            "Date", "String", "Object", "Variant", "User defined" };
        }

        private static bool HandleAsLong(String typeName)
        {
            string[] types = { "Integer", "Long", "Single" };
            return types.Contains(typeName);
        }

        private static bool HandleAsDouble(String typeName)
        {
            string[] types = { "Double","Double(negative)","Double(positive)", "Currency" };
            return types.Contains(typeName);
        }

        private List<IInspectionResult> InspectCaseClauses(QualifiedContext<ParserRuleContext> selectCaseStmt, List<QualifiedContext<ParserRuleContext>> caseClauses)
        {
            if (_typeName.Equals("String"))
            {
                return HandleSelectCase(selectCaseStmt, caseClauses, new List<Tuple<string, string, Func<string, string, string, string, bool>>>());
            }
            else if (HandleAsLong(_typeName))
            {
                return HandleSelectCase(selectCaseStmt, caseClauses, new List<Tuple<long, long, Func<long, long, long, long, bool>>>());
            }
            //else if (HandleAsDouble(_typeName))
            //{
            //    return HandleSelectCase(selectCaseStmt, caseClauses, new List<Tuple<double, double, Func<double, double, double, double, bool>>>());
            //}
            //else
            {
                return new List<IInspectionResult>();
            }
        }

        private List<IInspectionResult> HandleSelectCase(QualifiedContext<ParserRuleContext> selectCaseStmt, List<QualifiedContext<ParserRuleContext>> caseClauses, List<Tuple<long, long, Func<long, long, long, long, bool>>> clauseEvaluations)
        {
            var inspResults = new List<IInspectionResult>();

            foreach (var caseClause in caseClauses)
            {
                var rangeClauses = ParserRuleContextHelper.GetChildren<VBAParser.RangeClauseContext>(caseClause.Context as ParserRuleContext);
                foreach (var ctxt in rangeClauses)
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
                            bool addInspResult = false;
                            foreach (var cTest in clauseEvaluations)
                            {
                                if (cTest.Item3(startSelectValue, endSelectValue, cTest.Item1, cTest.Item2))
                                {
                                    addInspResult = true;
                                }
                            }
                            if (addInspResult)
                            {
                                inspResults = AddInspectionResult(caseClause, inspResults);
                            }
                            else
                            {
                                var clauseTest = new Tuple<long, long, Func<long, long, long, long, bool>>(startSelectValue, endSelectValue, _compareCaseEqualLongs);
                               // var xTest = new Tuple<long, long, Func<CompareParams<long>, bool>>(startSelectValue, endSelectValue, _compareCaseEqualLongs);
                                clauseEvaluations.Add(clauseTest);
                            }
                        }
                    }

                    var theOperator = "=";
                    //The 'Is' case
                    var opCtxt = ParserRuleContextHelper.GetChild<VBAParser.ComparisonOperatorContext>(ctxt);
                    if (opCtxt != null)
                    {
                        theOperator = opCtxt.GetText();
                    }

                    ////A statement like z > 5
                    //var relationalOpCtxt = ParserRuleContextHelper.GetChild<VBAParser.RelationalOpContext>(ctxt);
                    //if (relationalOpCtxt != null)
                    //{
                    //    theOperator = opCtxt.GetText();
                    //}

                    if (!_operatorsLong.ContainsKey(theOperator))
                    {
                        return inspResults;
                    }

                    var literalExpressionContexts = ParserRuleContextHelper.GetDescendents<VBAParser.LiteralExpressionContext>(ctxt as ParserRuleContext);
                    if (literalExpressionContexts.Any() && !hasStartEndContexts)
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
                                foreach (var cTest in clauseEvaluations)
                                {
                                    if (cTest.Item3(theValue, theValue, cTest.Item1, cTest.Item2))
                                    {
                                        addInspResult = true;
                                    }
                                }
                                if (addInspResult)
                                {
                                    inspResults = AddInspectionResult(caseClause, inspResults);
                                }
                                else
                                {
                                    Func<long, long, long, long, bool> comparison = _operatorsLong[theOperator];
                                    var clauseTest = new Tuple<long, long, Func<long, long, long, long, bool>>(theValue, theValue, comparison);
                                    clauseEvaluations.Add(clauseTest);
                                }
                            }
                        }
                    }
                }
            }
            return inspResults;
        }

        private List<IInspectionResult> HandleSelectCase(QualifiedContext<ParserRuleContext> selectCaseStmt, List<QualifiedContext<ParserRuleContext>> caseClauses, List<Tuple<string, string, Func<string, string, string, string, bool>>> clauseEvaluations)
        {
            var inspResults = new List<IInspectionResult>();

            foreach (var caseClause in caseClauses)
            {
                var rangeClauses = ParserRuleContextHelper.GetChildren<VBAParser.RangeClauseContext>(caseClause.Context as ParserRuleContext);
                foreach (var ctxt in rangeClauses)
                {
                    var literalExpressionContexts = ParserRuleContextHelper.GetDescendents<VBAParser.LiteralExpressionContext>(ctxt as ParserRuleContext);
                    foreach (var child in literalExpressionContexts)
                    {
                        var text = child.GetText();
                        bool addInspResult = false;
                        foreach (var stringClauseEval in clauseEvaluations)
                        {
                            foreach (var clauseEvaluation in clauseEvaluations)
                            {
                                if (clauseEvaluation.Item3(text, text, clauseEvaluation.Item1, clauseEvaluation.Item2))
                                {
                                    addInspResult = true;
                                }
                            }
                        }
                        if (addInspResult)
                        {
                            inspResults = AddInspectionResult(caseClause, inspResults);
                        }
                        else
                        {
                            //var xTest = new Tuple<string, string, Func<CompareParams<string>, bool>>(text, text, _compareCaseClauseStrings);
                            clauseEvaluations.Add(new Tuple<string, string, Func<string, string, string, string, bool>>(text, text, _compareCaseClauseStrings));
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

        private Func<string, string, string, string, bool> _compareCaseClauseStrings = delegate (string candidateMin, string candidateMax, string minVal, string maxVal)
        {
            //simple candidate value, simple value compare
            //if ((candidateMin.Equals(candidateMax)) && (minVal.Equals(maxVal)))
            //{
            return candidateMin.Equals(minVal);
            //}

            ////simple candidate value, range compare
            //if ((candidateMin == candidateMax) && (minVal != maxVal))
            //{
            //    Debug.Assert(maxVal >= minVal);
            //    return candidateMin >= minVal && candidateMin <= maxVal;
            //}

            ////candidate range, simple value compare
            //if ((candidateMin != candidateMax) && (minVal == maxVal))
            //{
            //    Debug.Assert(candidateMax >= candidateMin);
            //    return minVal >= candidateMin && maxVal <= candidateMax;
            //}

            ////range to range compare
            //Debug.Assert(candidateMax > candidateMin);
            //return (candidateMin >= minVal && candidateMin <= maxVal) || (candidateMax >= minVal && candidateMax <= maxVal);
        };

        private Func<long, long, long, long, bool> _compareCaseEqualLongs = delegate (long candidateMin, long candidateMax, long minVal, long maxVal)
        {
            //simple candidate value, simple value compare
            if ((candidateMin == candidateMax) && (minVal == maxVal))
            {
                return candidateMin == minVal;
            }

            //simple candidate value, range compare
            if ((candidateMin == candidateMax) && (minVal != maxVal))
            {
                Debug.Assert(maxVal >= minVal);
                return candidateMin >= minVal && candidateMin <= maxVal;
            }

            //candidate range, simple value compare
            if ((candidateMin != candidateMax) && (minVal == maxVal))
            {
                Debug.Assert(candidateMax >= candidateMin);
                return minVal >= candidateMin && maxVal <= candidateMax;
            }

            //range to range compare
            Debug.Assert(candidateMax > candidateMin);
            return (candidateMin >= minVal && candidateMin <= maxVal) || (candidateMax >= minVal && candidateMax <= maxVal);
        };

        private Func<long, long, long, long, bool> _compareCaseGTLongs = delegate (long candidateMin, long candidateMax, long minVal, long maxVal)
        {
            return HasValidParamsForSimpleCompare(candidateMin, candidateMax, minVal, maxVal) ? candidateMin.CompareTo(minVal) > 0 : false;
        };

        private Func<long, long, long, long, bool> _compareCaseLTLongs = delegate (long candidateMin, long candidateMax, long minVal, long maxVal)
        {
            return HasValidParamsForSimpleCompare(candidateMin, candidateMax, minVal, maxVal) ? candidateMin.CompareTo(minVal) < 0 : false;
        };

        private Func<long, long, long, long, bool> _compareCaseGTELongs = delegate (long candidateMin, long candidateMax, long minVal, long maxVal)
        {
            return HasValidParamsForSimpleCompare(candidateMin, candidateMax, minVal, maxVal) ? candidateMin.CompareTo(minVal) >= 0 : false;
        };

        private Func<long, long, long, long, bool> _compareCaseLTELongs = delegate (long candidateMin, long candidateMax, long minVal, long maxVal)
        {
            return HasValidParamsForSimpleCompare(candidateMin, candidateMax, minVal, maxVal) ? candidateMin.CompareTo(minVal) <= 0 : false; // candidateMin <= minVal : false;
        };

        private Func<long, long, long, long, bool> _compareCaseNEQLongs = delegate (long candidateMin, long candidateMax, long minVal, long maxVal)
        {
            return HasValidParamsForSimpleCompare(candidateMin, candidateMax, minVal, maxVal) ? candidateMin.CompareTo(minVal) != 0 : false;
        };

        private static bool HasValidParamsForSimpleCompare<T>( T candidateMin, T candidateMax, T minVal, T maxVal) where T : System.IComparable<T>
        {
            return (candidateMin.CompareTo(candidateMax) == 0);
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
