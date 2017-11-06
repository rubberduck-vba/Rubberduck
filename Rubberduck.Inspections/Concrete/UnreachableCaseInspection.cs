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
        private readonly string _unreachableCaseInspectionResultFormat = //"Unreachable or conflicting Case block";
        "The Select Case has already been handled by a previous case";
        //InspectionsUI.UnreachableCaseInspectionName;
        private readonly string _conflictingCaseInspectionResultFormat = //"Unreachable or conflicting Case block";
        "The Select Case has already been partially handled by a previous case";

        private readonly string _typeMismatchCaseInspectionResultFormat = //"Unreachable or conflicting Case block";
        "The Select Case is of a different Type than the Select Statement";

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
            SelectStmtWrapper selectStmt = new SelectStmtWrapper(State, selectCaseStmt);
            if (!selectStmt.TypeName.Equals(string.Empty))
            {
                return EvaluateCaseClauses(selectStmt);
            }

            return new List<ParserRuleContext>();
        }

        private List<ParserRuleContext> EvaluateCaseClauses(SelectStmtWrapper selectStmt)
        {
            var unreachableClauses = new List<ParserRuleContext>();
            var rangeEvals = LoadBoundaryValueEvaluations(selectStmt.TypeName, new List<IRangeClause>());

            foreach (var caseClause in selectStmt.CaseClauses)
            {
                unreachableClauses = EvaluateCaseClause(caseClause, rangeEvals, unreachableClauses);

                rangeEvals.AddRange(caseClause.RangeClauses.Where(rg => rg.IsParseable));
            }

            if (selectStmt.HasUnreachableCaseElse )
            {
                unreachableClauses.Add(selectStmt.CaseElseClause);
            }

            return unreachableClauses;
        }

        private List<ParserRuleContext> EvaluateCaseClause(CaseClauseWrapper caseClause, List<IRangeClause> priorRangeClauses, List<ParserRuleContext> unreachableClauses)
        {
            if (caseClause.IsUnreachable)
            {
                unreachableClauses.Add(caseClause.Value);
                return unreachableClauses;
            }

            var results = new List<RangeClauseComparer>();
            var selectCaseTypename = caseClause.Parent.TypeName;
            foreach (var rangeClause in caseClause.RangeClauses)
            {
                for (int idx = 0; idx < priorRangeClauses.Count(); idx++)
                {
                    var comp = new RangeClauseComparer();
                    comp.Compare(rangeClause, priorRangeClauses[idx], selectCaseTypename);
                    results.Add(comp);
                }
            }

            caseClause.Parent.HasUnreachableCaseElse = results.Where(result => result.CausesUnreachableCaseElse).Any();

            if (results.Where(result => !result.IsReachable).Any())
            {
                unreachableClauses.Add(caseClause.Value);
            }

            return unreachableClauses;
        }

        private static List<IRangeClause> LoadBoundaryValueEvaluations(string theTypeName, List<IRangeClause> rangeEvals)
        {
            //return rangeEvals;

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
            //private readonly VBAParser.SelectExpressionContext _selectExprContext;

            private IdentifierReference _idReference;
            private readonly List<CaseClauseWrapper> _caseClauses;
            private VBAParser.CaseElseClauseContext _caseElseClause;
            private bool _hasUnreachableCaseElse;

            //public bool HasValue => _selectStmtCtxt != null;
            //private QualifiedContext<ParserRuleContext> Value => _selectStmtCtxt;

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
                //_selectExprContext = ParserRuleContextHelper.GetChild<VBAParser.SelectExpressionContext>(_selectStmtCtxt.Context);
                _typeName = DetermineTheTypeName(SelectExprContext);

                _caseClauses = new List<CaseClauseWrapper>();
                foreach (var caseClause in ParserRuleContextHelper.GetChildren<VBAParser.CaseClauseContext>(_selectStmtCtxt.Context))
                {
                    _caseClauses.Add(new CaseClauseWrapper(this, caseClause));
                }
                _caseElseClause = ParserRuleContextHelper.GetChild<VBAParser.CaseElseClauseContext>(_selectStmtCtxt.Context);
            }

            public bool HasUnreachableCaseElse
            {
                set { if (CaseElseClause != null) _hasUnreachableCaseElse = value; else _hasUnreachableCaseElse = false; }
                get { return CaseElseClause != null ? _hasUnreachableCaseElse : false; }
            }

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

            public CaseClauseWrapper(SelectStmtWrapper selectStmt, VBAParser.CaseClauseContext caseClause)
            {
                _parent = selectStmt;
                _caseClause = caseClause;
                _rangeClauses = new List<RangeClause>();
                HasInternalConflict = false;
                var rangeClauseContexts = ParserRuleContextHelper.GetChildren<VBAParser.RangeClauseContext>(_caseClause);
                foreach (var rangeClauseCtxt in rangeClauseContexts)
                {
                    var rangeClause = new RangeClause(this, rangeClauseCtxt);
                    if (rangeClause.IsParseable)
                    {
                        _rangeClauses.Add(new RangeClause(this, rangeClauseCtxt));
                    }
                    else
                    {
                        if (!rangeClause.MatchesSelectCaseType)
                        {
                            IsUnreachable = true;
                        }
                    }
                }
            }

            public VBAParser.CaseClauseContext Value => _caseClause;
            public SelectStmtWrapper Parent => _parent;
            public List<RangeClause> RangeClauses => _rangeClauses;
            public bool HasInternalConflict { set; get; }
            public bool IsUnreachable { set; get; }
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
