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
    //TODO: Add replace with UI Resource
    public static class CaseInspectionMessages
    {
        public static string Unreachable => "Unreachable Case Statement: Handled by previous Case statement(s)";
        //public static string Overlap => "The Case is partially handled by a previous case";
        //public static string InternalConflict => "The Case statment has redundant or overlapping criteria";
        public static string Mismatch => "Unreachable Case Statement: Type does not match the Select Statement";
        public static string ExceedsBoundary => "Unreachable Case Statement: Value is not valid for the Select Statement Type";
        public static string CaseElse => "Unreachable Case Else Statement: All possible values are handled by previous Case statement(s)";
    }

    public sealed class SelectCaseInspection : ParseTreeInspectionBase
    {
        private struct PriorHandlers
        {
            public VBAValue GoverningIsLT;
            public VBAValue GoverningIsGT;
            public List<VBAValue> SingleValues;
            public List<Tuple<VBAValue, VBAValue>> Ranges;
            public List<string> Indeterminants;
        }

        public SelectCaseInspection(RubberduckParserState state)
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
                SelectStmtWrapper selectStmtWrapper = new SelectStmtWrapper(State, selectStmt);
                if (selectStmtWrapper.TypeName.Equals(string.Empty))
                {
                    continue;
                }

                var reportableCaseClauseResults = EvaluateSelectCaseClauses(selectStmtWrapper);

                foreach (var clauseWrapper in reportableCaseClauseResults)
                {
                    string msg = string.Empty;
                    if (clauseWrapper.IsHandledByPriorClause)
                    {
                        msg = CaseInspectionMessages.Unreachable;
                    }
                    else if (clauseWrapper.HasInconsistentType)
                    {
                        msg = CaseInspectionMessages.Mismatch;
                    }
                    else if (clauseWrapper.HasOutofRangeValue)
                    {
                        msg = CaseInspectionMessages.ExceedsBoundary;
                    }
                    else if(clauseWrapper.IsCaseElse && selectStmtWrapper.CaseElseIsUnreachable)
                    {
                        msg = CaseInspectionMessages.CaseElse;
                    }

                    if (!msg.Equals(string.Empty))
                    {
                        inspResults.Add(CreateInspectionResult(selectStmt, clauseWrapper.CaseContext, msg));
                    }
                }
            }
            return inspResults;
        }

        private IInspectionResult CreateInspectionResult(QualifiedContext<ParserRuleContext> selectStmt, ParserRuleContext unreachableBlock, string message)
        {
            return new QualifiedContextInspectionResult(this,
                        message,
                        new QualifiedContext<ParserRuleContext>(selectStmt.ModuleName, unreachableBlock));
        }

        private List<CaseClauseWrapper> EvaluateSelectCaseClauses(SelectStmtWrapper selectCaseStmt)
        {
            var currentHandlers = CheckBoundaries(selectCaseStmt);

            var reportableCaseClauseResults = new List<CaseClauseWrapper>();
            foreach (var caseClause in selectCaseStmt.CaseClauses)
            {
                caseClause.IsHandledByPriorClause = selectCaseStmt.CaseElseIsUnreachable;

                if (caseClause.HasInconsistentType
                    || caseClause.HasOutofRangeValue
                    || caseClause.IsHandledByPriorClause)
                {
                    reportableCaseClauseResults.Add(caseClause);
                    continue;
                }
                else
                {
                    currentHandlers = EvaluateCaseClause(caseClause, currentHandlers);
                    if (caseClause.IsHandledByPriorClause)
                    {
                        reportableCaseClauseResults.Add(caseClause);
                    }
                    if (!caseClause.Parent.CaseElseIsUnreachable)
                    {
                        caseClause.Parent.CaseElseIsUnreachable = caseClause.MakesRemainingClausesUnreachable;
                    }
                }
            }
            if (selectCaseStmt.CaseElseIsUnreachable)
            {
                reportableCaseClauseResults.Add(new CaseClauseWrapper(selectCaseStmt, selectCaseStmt.CaseElseClause));
            }
            return reportableCaseClauseResults;
        }

        private PriorHandlers EvaluateCaseClause(CaseClauseWrapper caseClause, PriorHandlers currentHandlers)
        {
            foreach (var range in caseClause.RangeClauses)
            {
                if (range.IsSingleVal)
                {
                    if (range.CompareByTextOnly)
                    {
                        if (currentHandlers.Indeterminants.Contains(range.Context.GetText()))
                        {
                            range.IsPreviouslyHandled = true;
                        }
                        else
                        {
                            currentHandlers.Indeterminants.Add(range.Context.GetText());
                        }
                        continue;
                    }

                    var theValue = new VBAValue(range.ValueAsString, caseClause.Parent.TypeName);
                    if (!range.UsesIsClause)
                    {
                        currentHandlers = HandleSimpleSingleValueCompare(range, theValue, currentHandlers);
                    }
                    else  //Uses 'Is' clauses
                    {
                        if (new string[] { CompareSymbols.LT, CompareSymbols.LTE, CompareSymbols.GT, CompareSymbols.GTE }.Contains(range.CompareSymbol))
                        {
                            currentHandlers = SetGoverningRelationChecks(theValue, range.CompareSymbol, currentHandlers);
                        }
                        else if (CompareSymbols.EQ.Equals(range.CompareSymbol))
                        {
                            currentHandlers = HandleSimpleSingleValueCompare(range, theValue, currentHandlers);
                        }
                        else if (CompareSymbols.NEQ.Equals(range.CompareSymbol))
                        {
                            if(currentHandlers.SingleValues.Contains(theValue))
                            {
                                range.CausesUnreachableCaseElse = true;
                            }

                            currentHandlers = SetGoverningRelationChecks(theValue, CompareSymbols.LT, currentHandlers);
                            currentHandlers = SetGoverningRelationChecks(theValue, CompareSymbols.GT, currentHandlers);
                        }
                    }
                }
                else  //It is a Case Statemnt like "Case 45 To 150"
                {
                    currentHandlers = AggregateRanges(currentHandlers);

                    var theMinValue = new VBAValue(range.ValueMinAsString, caseClause.Parent.TypeName);
                    var theMaxValue = new VBAValue(range.ValueMaxAsString, caseClause.Parent.TypeName);

                    range.IsPreviouslyHandled = currentHandlers.Ranges.Where(rg => theMinValue.IsWithin(rg.Item1, rg.Item2) 
                            && theMaxValue.IsWithin(rg.Item1, rg.Item2)).Any()
                            || currentHandlers.GoverningIsLT != null && currentHandlers.GoverningIsLT > theMaxValue
                            || currentHandlers.GoverningIsGT != null && currentHandlers.GoverningIsGT < theMinValue;

                    if (!range.IsPreviouslyHandled)
                    {
                        var overlapsMin = currentHandlers.Ranges.Where(rg => theMinValue.IsWithin(rg.Item1, rg.Item2));
                        var overlapsMax = currentHandlers.Ranges.Where(rg => theMaxValue.IsWithin(rg.Item1, rg.Item2));
                        var updated = new List<Tuple<VBAValue, VBAValue>>();
                        foreach (var rg in currentHandlers.Ranges)
                        {
                            if (overlapsMin.Contains(rg))
                            {
                                updated.Add(new Tuple<VBAValue, VBAValue>(rg.Item1, theMaxValue));
                            }
                            else if (overlapsMax.Contains(rg))
                            {
                                updated.Add(new Tuple<VBAValue, VBAValue>(theMinValue, rg.Item2));
                            }
                            else
                            {
                                updated.Add(rg);
                            }
                        }

                        if (!overlapsMin.Any() && !overlapsMax.Any())
                        {
                            updated.Add(new Tuple<VBAValue, VBAValue>(theMinValue, theMaxValue));
                        }
                        currentHandlers.Ranges = updated;
                    }

                    if (caseClause.Parent.TypeName.Equals(Tokens.Boolean))
                    {
                        range.CausesUnreachableCaseElse = theMinValue != theMaxValue;
                    }
                }
            }

            caseClause.MakesRemainingClausesUnreachable = caseClause.RangeClauses.Where(rg => rg.CausesUnreachableCaseElse).Any();

            if (!caseClause.MakesRemainingClausesUnreachable)
            {
                caseClause.MakesRemainingClausesUnreachable = IsClausesCoverAllValues(currentHandlers);
            }

            caseClause.IsHandledByPriorClause = caseClause.RangeClauses.Where(rg => rg.IsPreviouslyHandled).Count() == caseClause.RangeClauses.Count;
            return currentHandlers;
        }

        private bool IsClausesCoverAllValues(PriorHandlers currentHandlers)
        {
            if (currentHandlers.GoverningIsLT != null && currentHandlers.GoverningIsLT != null)
            {
                return currentHandlers.GoverningIsLT > currentHandlers.GoverningIsGT
                        || (currentHandlers.GoverningIsLT >= currentHandlers.GoverningIsGT 
                        && currentHandlers.SingleValues.Contains(currentHandlers.GoverningIsLT));
            }
            return false;
        }

        private bool SingleValueIsHandledPreviously(VBAValue theValue, PriorHandlers current)
        {
            return current.GoverningIsLT != null && theValue < current.GoverningIsLT
                || current.GoverningIsGT != null && theValue > current.GoverningIsGT
                || current.SingleValues.Contains(theValue)
                || current.Ranges.Where(rg => theValue.IsWithin(rg.Item1, rg.Item2)).Any();
        }

        private PriorHandlers SetGoverningRelationChecks(VBAValue theValue, string compareSymbol, PriorHandlers current)
        {
            if (!(new string[] { CompareSymbols.LT, CompareSymbols.LTE, CompareSymbols.GT, CompareSymbols.GTE }.Contains(compareSymbol)))
            {
                return current;
            }

            if (new string[] { CompareSymbols.LT, CompareSymbols.LTE }.Contains(compareSymbol))
            {
                current.GoverningIsLT = current.GoverningIsGT == null ? theValue
                    : current.GoverningIsLT < theValue ? theValue : current.GoverningIsLT;
            }
            else
            {
                current.GoverningIsGT = current.GoverningIsGT == null ? theValue 
                    : current.GoverningIsGT > theValue ? theValue : current.GoverningIsGT;
            }

            if (new string[] { CompareSymbols.LTE, CompareSymbols.GTE }.Contains(compareSymbol))
            {
                if (!current.SingleValues.Contains(theValue))
                {
                    current.SingleValues.Add(theValue);
                }
            }
            return current;
        }

        private PriorHandlers HandleSimpleSingleValueCompare(SelectCaseInspectionRangeClause range, VBAValue theValue,  PriorHandlers current)
        {
            range.IsPreviouslyHandled = SingleValueIsHandledPreviously(theValue, current);

            if (theValue.TargetTypeName.Equals(Tokens.Boolean))
            {
                range.CausesUnreachableCaseElse = current.SingleValues.Any()
                    && !current.SingleValues.Contains(theValue);
            }

            if (!range.IsPreviouslyHandled)
            {
                current.SingleValues.Add(theValue);
            }
            return current;
        }

        private PriorHandlers CheckBoundaries(SelectStmtWrapper selectCaseStmt)
        {
            var currentHandlers = new PriorHandlers
            {
                GoverningIsGT = VBAValue.CreateBoundaryMax(selectCaseStmt.TypeName),
                GoverningIsLT = VBAValue.CreateBoundaryMin(selectCaseStmt.TypeName),
                SingleValues = new List<VBAValue>(),
                Ranges = new List<Tuple<VBAValue, VBAValue>>(),
                Indeterminants = new List<string>()
            };

            var reportableCaseClauseResults = new List<CaseClauseWrapper>   ();

            if (currentHandlers.GoverningIsGT == null && currentHandlers.GoverningIsLT == null)
            {
                return currentHandlers;
            }

            foreach ( var caseClause in selectCaseStmt.CaseClauses)
            {
                foreach (var range in caseClause.RangeClauses)
                {
                    if (!range.IsParseable)
                    {
                        continue;
                    }

                    if (range.IsSingleVal)
                    {
                        var theValue = new VBAValue(range.ValueAsString, selectCaseStmt.TypeName);
                        if (new string[] { CompareSymbols.EQ }.Contains(range.CompareSymbol))
                        {
                            range.HasOutOfBoundsValue = theValue < currentHandlers.GoverningIsLT || theValue > currentHandlers.GoverningIsGT;
                        }
                        else if (new string[] { CompareSymbols.GT, CompareSymbols.GTE }.Contains(range.CompareSymbol))
                        {
                            range.HasOutOfBoundsValue = theValue > currentHandlers.GoverningIsGT;
                        }
                        else if (new string[] { CompareSymbols.LT, CompareSymbols.LTE }.Contains(range.CompareSymbol))
                        {
                            range.HasOutOfBoundsValue = theValue < currentHandlers.GoverningIsLT;
                        }
                    }
                    else
                    {
                        var theMinValue = new VBAValue(range.ValueMinAsString, selectCaseStmt.TypeName);
                        var theMaxValue = new VBAValue(range.ValueMaxAsString, selectCaseStmt.TypeName);

                        range.HasOutOfBoundsValue = theMinValue > currentHandlers.GoverningIsGT;
                        if (!range.HasOutOfBoundsValue)
                        {
                            range.HasOutOfBoundsValue = theMaxValue < currentHandlers.GoverningIsLT;
                        }
                    }

                }
                caseClause.HasOutofRangeValue = caseClause.RangeClauses.Where(rg => rg.HasOutOfBoundsValue).Count() == caseClause.RangeClauses.Count;
            }

            return currentHandlers;
        }

        private PriorHandlers AggregateRanges(PriorHandlers currentHandlers)
        {
            var startingRangeCount = currentHandlers.Ranges.Count;
            if (startingRangeCount > 1)
            {
                do
                {
                    startingRangeCount = currentHandlers.Ranges.Count();
                    currentHandlers.Ranges = AppendRanges(currentHandlers.Ranges);
                } while (currentHandlers.Ranges.Count() < startingRangeCount);
            }
            return currentHandlers;
        }

        private List<Tuple<VBAValue, VBAValue>> AppendRanges(List<Tuple<VBAValue, VBAValue>> ranges)
        {
            if (ranges.Count() <= 1)
            {
                return ranges;
            }

            if (!ranges.First().Item1.IsIntegerNumber)
            {
                return ranges;
            }

            var updated = new List<Tuple<VBAValue, VBAValue>>();
            var combinedLastRange = false;

            for (var idx = 0; idx < ranges.Count(); idx++)
            {
                if (idx + 1 >= ranges.Count())
                {
                    if (!combinedLastRange)
                    {
                        updated.Add(ranges[idx]);
                    }
                    continue;
                }
                combinedLastRange = false;
                var theMin = ranges[idx].Item1;
                var theMax = ranges[idx].Item2;
                var theNextMin = ranges[idx+1].Item1;
                var theNextMax = ranges[idx+1].Item2;
                if(theMax.AsLong() == theNextMin.AsLong() - 1)
                {
                    updated.Add(new Tuple<VBAValue, VBAValue>(theMin,theNextMax));
                    combinedLastRange = true;
                }
                else if(theMin.AsLong() == theNextMax.AsLong() + 1)
                {
                    updated.Add(new Tuple<VBAValue, VBAValue>(theNextMin, theMax));
                    combinedLastRange = true;
                }
                else
                {
                    updated.Add(ranges[idx]);
                }
            }

            return updated;
        }

        #region SelectStmtWrapper

        internal class SelectStmtWrapper
        {
            private readonly QualifiedContext<ParserRuleContext> _selectStmtCtxt;
            private readonly RubberduckParserState _state;
            private readonly string _typeName;

            private IdentifierReference _idReference;
            private List<CaseClauseWrapper> _caseClauses;
            private VBAParser.CaseElseClauseContext _caseElseClause;
            private bool _hasUnreachableCaseElse;

            public RubberduckParserState State => _state;
            private VBAParser.SelectExpressionContext SelectExprContext => ParserRuleContextHelper.GetChild<VBAParser.SelectExpressionContext>(_selectStmtCtxt.Context);
            public IdentifierReference IdReference => _idReference;
            public string TypeName => _typeName;
            public List<CaseClauseWrapper> CaseClauses => _caseClauses;
            public VBAParser.CaseElseClauseContext CaseElseClause => _caseElseClause;

            public SelectStmtWrapper(RubberduckParserState state, QualifiedContext<ParserRuleContext> selectStmtCtxt)
            {
                _state = state;
                _selectStmtCtxt = selectStmtCtxt;
                _hasUnreachableCaseElse = false;
                _typeName = DetermineTheTypeName(SelectExprContext);

                var caseClauseCtxts = ParserRuleContextHelper.GetChildren<VBAParser.CaseClauseContext>(_selectStmtCtxt.Context);
                _caseClauses = caseClauseCtxts.Select(cc => new CaseClauseWrapper(this, cc)).ToList();
                _caseElseClause = ParserRuleContextHelper.GetChild<VBAParser.CaseElseClauseContext>(_selectStmtCtxt.Context);
            }

            public bool CaseElseIsUnreachable
            {
                set { if (!_hasUnreachableCaseElse) _hasUnreachableCaseElse = value;  }
                get { return  _hasUnreachableCaseElse; }
            }

            public bool HasCaseElse => CaseElseClause != null;

            private string DetermineTheTypeName(VBAParser.SelectExpressionContext selectExprCtxt)
            {
                var relationalOpCtxt = ParserRuleContextHelper.GetChild<VBAParser.RelationalOpContext>(selectExprCtxt);
                if (relationalOpCtxt != null)
                {
                    return Tokens.Boolean;
                }
                else
                {
                    var smplName = ParserRuleContextHelper.GetDescendent<VBAParser.SimpleNameExprContext>(selectExprCtxt);
                    if (smplName != null)
                    {
                        _idReference = GetTheSelectCaseReference(_selectStmtCtxt, smplName.GetText());
                        if (_idReference != null)
                        {
                            return _idReference.Declaration.AsTypeName;
                        }
                    }
                }
                return string.Empty;
            }

            private IdentifierReference GetTheSelectCaseReference(QualifiedContext<ParserRuleContext> selectCaseStmt, string theName)
            {
                var allRefs = new List<IdentifierReference>();
                foreach (var dec in _state.DeclarationFinder.MatchName(theName))
                {
                    allRefs.AddRange(dec.References);
                }

                if (!allRefs.Any())
                {
                    return null;
                }

                var selectCaseExpr = ParserRuleContextHelper.GetChild<VBAParser.SelectExpressionContext>(selectCaseStmt.Context);
                var selectCaseReference = allRefs.Where(rf => ParserRuleContextHelper.HasParent(rf.Context, selectCaseStmt.Context)
                                        && (ParserRuleContextHelper.HasParent(rf.Context, selectCaseExpr)));

                Debug.Assert(selectCaseReference.Count() == 1);
                return selectCaseReference.First();
            }
        }
        #endregion

#region CaseClauseWrapper
        internal class CaseClauseWrapper
        {
            private ParserRuleContext _clauseCtxt;
            private SelectStmtWrapper _parent;
            private List<SelectCaseInspectionRangeClause> _rangeClauses;
            private bool _isFullyEquivalent;
            private bool _isPartiallyEquivalent;
            public CaseClauseWrapper(SelectStmtWrapper selectStmt, VBAParser.CaseClauseContext caseClause)
            {
                _parent = selectStmt;
                _clauseCtxt = caseClause;
                _rangeClauses = new List<SelectCaseInspectionRangeClause>();
                HasInternalConflict = false;
                foreach (var rangeClauseCtxt in RangeClauseContexts)
                {
#if (DEBUG)
                    var test = rangeClauseCtxt.GetText();
#endif
                    var rangeClause = new SelectCaseInspectionRangeClause(this, rangeClauseCtxt);
                    if (!rangeClause.IsParseable)
                    {
                        if (!rangeClause.MatchesSelectCaseType)
                        {
                            HasInconsistentType = true;
                        }
                    }
                    _rangeClauses.Add(new SelectCaseInspectionRangeClause(this, rangeClauseCtxt));
                }
            }

            public CaseClauseWrapper(SelectStmtWrapper selectStmt, VBAParser.CaseElseClauseContext caseElse)
            {
                _parent = selectStmt;
                _clauseCtxt = caseElse;
                _rangeClauses = new List<SelectCaseInspectionRangeClause>();
                HasInternalConflict = false;
            }

            public ParserRuleContext CaseContext => _clauseCtxt;
            public SelectStmtWrapper Parent => _parent;
            public List<SelectCaseInspectionRangeClause> RangeClauses => _rangeClauses;
            public List<VBAParser.RangeClauseContext> RangeClauseContexts => ParserRuleContextHelper.GetChildren<VBAParser.RangeClauseContext>(_clauseCtxt);
            public bool HasInternalConflict { set; get; }
            public bool IsCaseElse => _clauseCtxt is VBAParser.CaseElseClauseContext;
            public bool IsHandledByPriorClause
            {
                set
                {
                    _isFullyEquivalent = value;
                    if (_isFullyEquivalent) { OverlapsWithPriorClause = false; }
                }
                get
                {
                    return _isFullyEquivalent;
                }
            }
            public bool OverlapsWithPriorClause
            {
                set
                {
                    if (!IsHandledByPriorClause) { _isPartiallyEquivalent = value; } else { _isPartiallyEquivalent = false; }
                }
                get
                {
                    return _isPartiallyEquivalent;
                }
            }
            public bool HasInconsistentType { set; get; }
            public bool HasOutofRangeValue { set; get; }
            public bool MakesRemainingClausesUnreachable { set; get; }
        }

        #endregion

        #region Listener
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
