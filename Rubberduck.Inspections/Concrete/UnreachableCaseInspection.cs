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
        public static string Unreachable => "The Select Case is unreachable or handled by a previous case";
        public static string Conflict => "The Select Case has already been partially handled by a previous case";
        public static string InternalConflict => "The Select Case contains redundant or overlapping criteria";
        public static string Mismatch => "The Select Case is of a different Type than the Select Statement";
        public static string CaseElse => "The Case Else clause is unreachable";
    }

    public sealed class UnreachableCaseInspection : ParseTreeInspectionBase
    {
        private readonly string _unreachableCaseInspectionResultFormat
        = CaseInspectionMessages.Unreachable;

        private readonly string _conflictingCaseInspectionResultFormat
        = CaseInspectionMessages.Conflict;

        private readonly string _internalClauseConflictFormat
        = CaseInspectionMessages.InternalConflict;

        private readonly string _typeMismatchCaseInspectionResultFormat
        = CaseInspectionMessages.Mismatch;

        private readonly string _unreachableCaseElse
        = CaseInspectionMessages.CaseElse;

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
                SelectStmtWrapper wrapper = new SelectStmtWrapper(State, selectStmt);
                if (wrapper.TypeName.Equals(string.Empty))
                {
                    continue;
                }

                var caseClauseResults = EvaluateSelectCaseClauses(wrapper);

                foreach(var clauseWrapper in caseClauseResults.Item2)
                {
                    string msg = string.Empty;
                    if (clauseWrapper.IsFullyHandled)
                    {
                        msg = _unreachableCaseInspectionResultFormat;
                    }
                    else if (clauseWrapper.IsPartiallyHandled)
                    {
                        msg = _conflictingCaseInspectionResultFormat;
                    }
                    else if (clauseWrapper.HasInternalConflict)
                    {
                        msg = _internalClauseConflictFormat;
                    }
                    else if (clauseWrapper.HasInconsistentType)
                    {
                        msg = _typeMismatchCaseInspectionResultFormat;
                    }
                    inspResults.Add(CreateInspectionResult(selectStmt, clauseWrapper.Value, msg));
                }
                if (caseClauseResults.Item1 != null)
                {
                    inspResults.Add(CreateInspectionResult(selectStmt, caseClauseResults.Item1, _unreachableCaseElse));
                }
            }
            return inspResults;
        }

        private IInspectionResult CreateInspectionResult(QualifiedContext<ParserRuleContext> selectStmt, ParserRuleContext unreachableBlock)
        {
            return new QualifiedContextInspectionResult(this,
                        _unreachableCaseInspectionResultFormat,
                        new QualifiedContext<ParserRuleContext>(selectStmt.ModuleName, unreachableBlock));
        }

        private IInspectionResult CreateInspectionResult(QualifiedContext<ParserRuleContext> selectStmt, ParserRuleContext unreachableBlock, string message)
        {
            return new QualifiedContextInspectionResult(this,
                        message,
                        new QualifiedContext<ParserRuleContext>(selectStmt.ModuleName, unreachableBlock));
        }

        private Tuple<VBAParser.CaseElseClauseContext,List<CaseClauseWrapper>> EvaluateSelectCaseClauses(SelectStmtWrapper selectCaseStmt)
        {
            Debug.Assert(!selectCaseStmt.TypeName.Equals(string.Empty));

            var caseClauseEvaluationResults = new List<CaseClauseWrapper>();
            var priorRangeClauses = new List<IRangeClause>();
            foreach (var caseClause in selectCaseStmt.CaseClauses)
            {
                if (caseClause.HasInconsistentType)
                {
                    caseClauseEvaluationResults.Add(caseClause);
                }
                else if (selectCaseStmt.CaseElseIsUnreachable)
                {
                    caseClause.IsFullyHandled = true;
                    caseClauseEvaluationResults.Add(caseClause);
                }
                else
                {
                    caseClauseEvaluationResults = EvaluateCaseClauseAgainstRangeExtents(caseClause, caseClauseEvaluationResults);
                    caseClauseEvaluationResults = EvaluateCaseClauseForInternalConflicts(caseClause, caseClauseEvaluationResults);
                    caseClauseEvaluationResults = EvaluateCaseClauseAgainstPriorRangeContexts(caseClause, priorRangeClauses, caseClauseEvaluationResults);
                    priorRangeClauses.AddRange(caseClause.RangeClauses.Where(rg => rg.MatchesSelectCaseType));
                }
            }

            return new Tuple<VBAParser.CaseElseClauseContext, List<CaseClauseWrapper>>
            (
                selectCaseStmt.CaseElseIsUnreachable ? selectCaseStmt.CaseElseClause : null,
                caseClauseEvaluationResults
            ); 
        }

        private List<CaseClauseWrapper> EvaluateCaseClauseAgainstRangeExtents(CaseClauseWrapper caseClause, List<CaseClauseWrapper> unreachableClauses)
        {
            var rangeExtents = LoadBoundaryValueEvaluations(caseClause.Parent.TypeName, new List<IRangeClause>());
            var allResults = new List<RangeClauseComparer.CompareResultData>();

            var comparer = new RangeClauseComparer();
            foreach (var rangeClause in caseClause.RangeClauses)
            {
                var text = rangeClause.Context.GetText();


                for (int idx = 0; idx < rangeExtents.Count(); idx++)
                {
                    var compResult = new RangeClauseComparer.CompareResultData();
                    if (!rangeClause.CompareByTextOnly)
                    {
                        compResult = comparer.Compare(rangeClause, rangeExtents[idx], caseClause.Parent.TypeName);
                    }
                    allResults.Add(compResult);
                }

                caseClause.IsFullyHandled = allResults.Where(res => res.IsRedundant).Any();
                if (caseClause.IsFullyHandled)
                {
                    unreachableClauses.Add(caseClause);
                    return unreachableClauses;
                }
            }
            return unreachableClauses;
        }


        private List<CaseClauseWrapper> EvaluateCaseClauseAgainstPriorRangeContexts(CaseClauseWrapper caseClause, List<IRangeClause> priorRangeClauses, List<CaseClauseWrapper> unreachableClauses)
        {
            if (caseClause.IsFullyHandled)
            {
                return unreachableClauses;
            }

            var allResults = new List<RangeClauseComparer.CompareResultData>();

            var comparer = new RangeClauseComparer();
            foreach (var rangeClause in caseClause.RangeClauses)
            {
                var text = rangeClause.Context.GetText();

                for (int idx = 0; idx < priorRangeClauses.Count(); idx++)
                {
                    var compResult = new RangeClauseComparer.CompareResultData();
                    if (priorRangeClauses[idx] is RangeClause)
                    {
                        var priorRgClause = (RangeClause)priorRangeClauses[idx];
                        var priorText = priorRgClause.Context.GetText();
                        if (rangeClause.CompareByTextOnly || priorRgClause.CompareByTextOnly)
                        {
                            if (rangeClause.Context.GetText().Equals(priorRgClause.Context.GetText()))
                            {
                                compResult.IsRedundant = true;
                            }
                        }
                        else
                        {
                            compResult = comparer.Compare(rangeClause, priorRangeClauses[idx], caseClause.Parent.TypeName);
                        }
                    }
                    allResults.Add(compResult);
                }

                caseClause.IsFullyHandled = allResults.Where(res => res.IsRedundant).Any();
                caseClause.IsPartiallyHandled = allResults.Where(res => res.HasInternalConflict).Any();
            }

            if (!caseClause.Parent.CaseElseIsUnreachable)
            {
                caseClause.Parent.CaseElseIsUnreachable = allResults.Where(result => result.MakesAllRemainingCasesUnreachable).Any();
            }

            if (caseClause.IsFullyHandled || caseClause.IsPartiallyHandled)
            {
                unreachableClauses.Add(caseClause);
            }

            return unreachableClauses;
        }

        private List<CaseClauseWrapper> EvaluateCaseClauseForInternalConflicts(CaseClauseWrapper caseClause, List<CaseClauseWrapper> unreachableClauses)
        {
            var allResults = new List<RangeClauseComparer.CompareResultData>();

            var comparer = new RangeClauseComparer();
            for (var currentIdx = caseClause.RangeClauses.Count-1; currentIdx > 0; currentIdx--) // var rangeClause in caseClause.RangeClauses)
            {
                var rangeClause = caseClause.RangeClauses[currentIdx];
                var text = rangeClause.Context.GetText();
                var priorRangeClauses = new List<IRangeClause>();
                for (var idxPrior = 0; idxPrior < currentIdx; idxPrior++)
                {
                    priorRangeClauses.Add(caseClause.RangeClauses[idxPrior]);
                }
                for(var idx = 0; idx < priorRangeClauses.Count; idx++)
                {
                    var compResult = new RangeClauseComparer.CompareResultData();
                    var priorRgClause = (RangeClause)priorRangeClauses[idx]; // caseClause.RangeClauses[idx];
                        var priorText = priorRgClause.Context.GetText();
                        if (rangeClause.CompareByTextOnly || priorRgClause.CompareByTextOnly)
                        {
                            if (rangeClause.Context.GetText().Equals(priorRgClause.Context.GetText()))
                            {
                                compResult.IsRedundant = true;
                            }
                        }
                        else
                        {
                            compResult = comparer.Compare(rangeClause, priorRgClause, caseClause.Parent.TypeName);
                        }
                    allResults.Add(compResult);
                }

                if (!caseClause.HasInternalConflict)
                {
                    caseClause.HasInternalConflict = allResults.Where(res => res.IsRedundant).Any();
                    if (!caseClause.HasInternalConflict)
                    {
                        caseClause.HasInternalConflict = allResults.Where(res => res.HasInternalConflict).Any();
                    }
                }
            }

            if (caseClause.HasInternalConflict)
            {
                unreachableClauses.Add(caseClause);
            }
            if (!caseClause.Parent.CaseElseIsUnreachable)
            {
                caseClause.Parent.CaseElseIsUnreachable = allResults.Where(result => result.MakesAllRemainingCasesUnreachable).Any();
            }

            return unreachableClauses;
        }

        private static List<IRangeClause> LoadBoundaryValueEvaluations(string theTypeName, List<IRangeClause> rangeEvals)
        {
            if (theTypeName.Equals(Tokens.Long))
            {
                long LONGMIN = -2147486648;
                long LONGMAX = 2147486647;
                return LoadRangeExtents(LONGMIN, LONGMAX, rangeEvals);
            }
            else if (theTypeName.Equals(Tokens.Integer))
            {
                long INTEGERMIN = -32768;
                long INTEGERMAX = 32767;
                return LoadRangeExtents(INTEGERMIN, INTEGERMAX, rangeEvals);
            }
            else if (theTypeName.Equals(Tokens.Byte))
            {
                long BYTEMIN = 0;
                long BYTEMAX = 255;
                return LoadRangeExtents(BYTEMIN, BYTEMAX, rangeEvals);
            }
            else if (theTypeName.Equals(Tokens.Boolean))
            {
                //If a value can be parsed to a Boolean, it's in range.
                return rangeEvals;
            }
            else if (theTypeName.Equals(Tokens.Currency))
            {
                decimal CURRENCYMIN = -922337203685477.5808M;
                decimal CURRENCYMAX = 922337203685477.5807M;
                return LoadRangeExtents(CURRENCYMIN, CURRENCYMAX, rangeEvals);
            }
            else if (theTypeName.Equals(Tokens.Single))
            {
                double SINGLEMIN = -3402823E38;
                double SINGLEMAX = 3402823E38;
                return LoadRangeExtents(SINGLEMIN, SINGLEMAX, rangeEvals);
            }
            //Decimal/Variant data type not supported by extent checks

            return rangeEvals;
        }

        private static List<IRangeClause> LoadRangeExtents<T>(T min, T max, List<IRangeClause> rangeEvals) where T : System.IComparable
        {
            rangeEvals.Add(new RangeClauseExtent<T>(min, "<"));
            rangeEvals.Add(new RangeClauseExtent<T>(max, ">"));
            return rangeEvals;
        }
        #region SelectStmtWrapper

        internal class SelectStmtWrapper
        {
            private readonly QualifiedContext<ParserRuleContext> _selectStmtCtxt;
            private readonly RubberduckParserState _state;
            private readonly string _typeName;

            private IdentifierReference _idReference;
            private readonly IEnumerable<CaseClauseWrapper> _caseClauses;
            private VBAParser.CaseElseClauseContext _caseElseClause;
            private bool _hasUnreachableCaseElse;

            public RubberduckParserState State => _state;
            private VBAParser.SelectExpressionContext SelectExprContext => ParserRuleContextHelper.GetChild<VBAParser.SelectExpressionContext>(_selectStmtCtxt.Context);
            public IdentifierReference IdReference => _idReference;
            public string TypeName => _typeName;
            public IEnumerable<CaseClauseWrapper> CaseClauses => _caseClauses;
            public VBAParser.CaseElseClauseContext CaseElseClause => _caseElseClause;

            public SelectStmtWrapper(RubberduckParserState state, QualifiedContext<ParserRuleContext> selectStmtCtxt)
            {
                _state = state;
                _selectStmtCtxt = selectStmtCtxt;
                _hasUnreachableCaseElse = false;
                _typeName = DetermineTheTypeName(SelectExprContext);

                var caseClauseCtxts = ParserRuleContextHelper.GetChildren<VBAParser.CaseClauseContext>(_selectStmtCtxt.Context);
                _caseClauses = caseClauseCtxts.Select(cc => new CaseClauseWrapper(this, cc));
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
            private VBAParser.CaseClauseContext _caseClause;
            private SelectStmtWrapper _parent;
            private List<RangeClause> _rangeClauses;
            private bool _isFullyEq;
            private bool _isPartiallyEq;

            public CaseClauseWrapper(SelectStmtWrapper selectStmt, VBAParser.CaseClauseContext caseClause)
            {
                _parent = selectStmt;
                _caseClause = caseClause;
                _rangeClauses = new List<RangeClause>();
                HasInternalConflict = false;
                foreach (var rangeClauseCtxt in RangeClauseContexts)
                {
                    var test = rangeClauseCtxt.GetText();
                    var rangeClause = new RangeClause(this, rangeClauseCtxt);
                    if (!rangeClause.IsParseable)
                    {
                        if (!rangeClause.MatchesSelectCaseType)
                        {
                            HasInconsistentType = true;
                            IsUnreachable = true;
                        }
                    }
                    _rangeClauses.Add(new RangeClause(this, rangeClauseCtxt));
                }
            }

            public VBAParser.CaseClauseContext Value => _caseClause;
            public SelectStmtWrapper Parent => _parent;
            public List<RangeClause> RangeClauses => _rangeClauses;
            public List<VBAParser.RangeClauseContext> RangeClauseContexts => ParserRuleContextHelper.GetChildren<VBAParser.RangeClauseContext>(_caseClause);
            public bool HasInternalConflict { set; get; }
            public bool IsUnreachable { set; get; }
            public bool IsFullyHandled
            {
                set
                {
                    _isFullyEq = value;
                    if (_isFullyEq) { IsPartiallyHandled = false; }
                }
                get
                {
                    return _isFullyEq;
                }
            }
            public bool IsPartiallyHandled
            {
                set
                {
                    if (!IsFullyHandled) { _isPartiallyEq = value; } else { _isPartiallyEq = false; }
                }
                get
                {
                    return _isPartiallyEq;
                }
            }
            public bool HasInconsistentType { set; get; }
            public bool CausesUnreachableForRemainingClauses { set; get; }
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
