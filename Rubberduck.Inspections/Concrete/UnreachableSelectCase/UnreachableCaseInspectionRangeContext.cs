using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Antlr4.Runtime;
using Antlr4.Runtime.Misc;
using Antlr4.Runtime.Tree;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Inspections.Abstract;

namespace Rubberduck.Inspections.Concrete.UnreachableSelectCase
{
    public interface IUnreachableCaseInspectionContext
    {
        ISummaryCoverage SummaryCoverage { get; }
        string EvaluationTypeName { set; get; }
    }

    public interface IUnreachableCaseInspectionRange : IUnreachableCaseInspectionContext
    {
        bool HasIncompatibleType { set; get; }
        bool IsValueRange { get; }
        bool IsLTorGT { get; }
        bool IsSingleValue { get; }
        bool IsRelationalOp { get; }
        ParserRuleContext Result { get; }
        ParserRuleContext RangeStart { get; }
        ParserRuleContext RangeEnd { get; }
        string IsClauseSymbol { get; }
        ISummaryCoverage Coverage(IUnreachableCaseInspectionValueResults results, string evaluationTypeName);
    }

    public abstract class UnreachableCaseInspectionContext
    {
        protected readonly ParserRuleContext _context;
        protected readonly IUnreachableCaseInspectionValueResults _inspValues;
        protected readonly IUnreachableCaseInspectionSummaryClauseFactory _factorySummaryClause;
        protected readonly IUnreachableCaseInspectionValueFactory _factoryValue;

        protected ISummaryCoverage _summaryCoverage;
        public UnreachableCaseInspectionContext(ParserRuleContext context, IUnreachableCaseInspectionValueResults inspValues, IUnreachableCaseInspectionSummaryClauseFactory factory, IUnreachableCaseInspectionValueFactory valueFactory)
        {
            _context = context;
            _inspValues = inspValues;
            _factorySummaryClause = factory;
            _factoryValue = valueFactory;
        }

        protected abstract bool IsResultContext<TContext>(TContext context) where TContext : ParserRuleContext;
        public ISummaryCoverage SummaryCoverage => _summaryCoverage;

        public TContext GetChild<TContext>() where TContext : ParserRuleContext
        {
            return Context.GetChild<TContext>();
        }

        protected ParserRuleContext Context => _context;

        protected static bool TryDetermineEvaluationTypeFromTypes(IEnumerable<string> typeNames, out string typeName)
        {
            typeName = string.Empty;
            var typeList = typeNames.ToList();

            //If everything is declared as a Variant, we do not attempt to inspect the selectStatement
            if (typeList.All(tn => new string[] { Tokens.Variant }.Contains(tn)))
            {
                return false;
            }
            typeList.All(tn => new string[] { typeList.First() }.Contains(tn));
            //If all match, the typeName is easy...This is the only way to return "String" or "Currency".
            if (typeList.All(tn => new string[] { typeList.First() }.Contains(tn)))
            {
                typeName = typeList.First();
                return true;
            }
            //Integer numbers will be evaluated using Long
            if (typeList.All(tn => new string[] { Tokens.Long, Tokens.Integer, Tokens.Byte }.Contains(tn)))
            {
                typeName = Tokens.Long;
                return true;
            }

            //Mix of Integertypes and rational number types will be evaluated using Double
            if (typeList.All(tn => new string[] { Tokens.Long, Tokens.Integer, Tokens.Byte, Tokens.Single, Tokens.Double }.Contains(tn)))
            {
                typeName = Tokens.Double;
                return true;
            }

            return false;
        }
    }

    public interface IUnreachableCaseInspectionSelectStmt : IUnreachableCaseInspectionContext
    {
        bool CanBeInspected { get; }
        void EvaluateForUnreachableCases();
        List<ParserRuleContext> UnreachableCases { get; }
        List<ParserRuleContext> MismatchTypeCases { get; }
        List<ParserRuleContext> UnreachableCaseElseCases { get; }
    }

    public class UnreachableCaseInspectionSelectStmt : UnreachableCaseInspectionContext, IUnreachableCaseInspectionSelectStmt
    {
        private readonly VBAParser.SelectCaseStmtContext _selectCaseContext;
        private List<ParserRuleContext> _unreachable;
        private List<ParserRuleContext> _misMatch;
        private List<ParserRuleContext> _caseElse;
        private string _evalTypeName;

        public UnreachableCaseInspectionSelectStmt(VBAParser.SelectCaseStmtContext selectCaseContext, IUnreachableCaseInspectionValueResults inspValues, IUnreachableCaseInspectionSummaryClauseFactory factory, IUnreachableCaseInspectionValueFactory valueFactory) 
            : base(selectCaseContext, inspValues, factory, valueFactory)
        {
            _selectCaseContext = selectCaseContext;
            _unreachable = new List<ParserRuleContext>();
            _misMatch = new List<ParserRuleContext>();
            _caseElse = new List<ParserRuleContext>();
        }

        public string EvaluationTypeName
        {
            set { }
            get => _evalTypeName ?? DetermineSelectCaseEvaluationTypeName(Context, _inspValues, _factorySummaryClause);
        }
        public bool CanBeInspected => !(EvaluationTypeName is null  || EvaluationTypeName.Equals(Tokens.Variant));
        public List<ParserRuleContext> UnreachableCases => _unreachable;
        public List<ParserRuleContext> MismatchTypeCases => _misMatch;
        public List<ParserRuleContext> UnreachableCaseElseCases => _caseElse;
        protected override bool IsResultContext<TContext>(TContext context)
        {
            return context is VBAParser.SelectExpressionContext;
        }

        public void EvaluateForUnreachableCases()
        {
            if (!CanBeInspected)
            {
                return;
            }

            _summaryCoverage = _factorySummaryClause.Create(EvaluationTypeName, _factoryValue);
            foreach (var caseClause in _selectCaseContext.caseClause())
            {
                if (_summaryCoverage.CoversAllValues)
                {
                    //Once all possble values are covered, the remaining CaseClauses are unreachable
                    _unreachable.Add(caseClause);
                    continue;
                }

                var caseClauseCoverge = _factorySummaryClause.Create(EvaluationTypeName, _factoryValue);
                var caseRanges = new List<IUnreachableCaseInspectionRange>();
                foreach (var range in caseClause.rangeClause())
                {
                    IUnreachableCaseInspectionRange inspRange = new UnreachableCaseInspectionRange(range, _inspValues, _factorySummaryClause, _factoryValue)
                    {
                        EvaluationTypeName = _evalTypeName
                    };
                    caseRanges.Add(inspRange);
                    caseClauseCoverge.Add(inspRange.Coverage(_inspValues, _summaryCoverage.TypeName));
                }

                //var difference = caseClauseCoverge.GetDifference(_summaryCoverage);
                var filterResults = _factorySummaryClause.Create(caseClauseCoverge.TypeName, _factoryValue);
                //if (difference.HasCoverage)
                if (caseClauseCoverge.TryFilterOutRedundateClauses(_summaryCoverage, ref filterResults))
                {
                    _summaryCoverage.Add(filterResults);
                }
                else //the caseClause contributes no new coverage, it is unreachable for one or two reasons
                {
                    if (caseRanges.All(cr => cr.HasIncompatibleType))
                    {
                        //Calling out CaseClauses that cannot be implicitly converted to the SelectCase type as a special case of unreachable
                        _misMatch.Add(caseClause);
                    }
                    else
                    {
                        _unreachable.Add(caseClause);
                    }
                }
            }
            if (_summaryCoverage.CoversAllValues && !(_selectCaseContext.caseElseClause() is null))
            {
                _caseElse.Add(_selectCaseContext.caseElseClause());
            }
        }

        private string DetermineSelectCaseEvaluationTypeName(ParserRuleContext context, IUnreachableCaseInspectionValueResults inspValues, IUnreachableCaseInspectionSummaryClauseFactory factory)
        {
            var selectStmt = (VBAParser.SelectCaseStmtContext)context;
            if (TryDetectTypeHint(selectStmt.selectExpression().GetText(), out _evalTypeName))
            {
                return _evalTypeName;
            }

            var typeName = string.Empty;
            if (inspValues.TryGetValue(selectStmt.selectExpression(), out IUnreachableCaseInspectionValue result))
            {
                _evalTypeName = result.TypeName;
            }

            if (TypeNameCandBeInspected(_evalTypeName))
            {
                return _evalTypeName;
            }
            return DeriveTypeFromCaseClauses(inspValues, selectStmt, factory);
        }

        private static bool TypeNameCandBeInspected(string typeName) => !(typeName == string.Empty || typeName == Tokens.Variant);

        private string DeriveTypeFromCaseClauses(IUnreachableCaseInspectionValueResults inspValues, VBAParser.SelectCaseStmtContext selectStmt, IUnreachableCaseInspectionSummaryClauseFactory factory)
        {
            var caseClauseTypeNames = new List<string>();
            foreach (var caseClause in selectStmt.caseClause())
            {
                foreach (var range in caseClause.rangeClause())
                {
                    if (TryDetectTypeHint(range.GetText(), out string hintTypeName))
                    {
                        caseClauseTypeNames.Add(hintTypeName);
                        continue;
                    }

                    var inspRange = new UnreachableCaseInspectionRange(range, inspValues, factory, _factoryValue);
                    if (inspRange.IsValueRange)
                    {
                        caseClauseTypeNames.Add(inspValues.GetTypeName(inspRange.RangeStart));
                        caseClauseTypeNames.Add(inspValues.GetTypeName(inspRange.RangeEnd));
                    }
                    else if (inspRange.IsRelationalOp)
                    {
                        caseClauseTypeNames.Add(Tokens.Boolean);
                    }
                    else
                    {
                        caseClauseTypeNames.Add(inspValues.GetTypeName(inspRange.Result));
                    }
                }
            }

            if (TryDetermineEvaluationTypeFromTypes(caseClauseTypeNames, out _evalTypeName))
            {
                return _evalTypeName;
            }
            return string.Empty;
        }

        private static bool TryDetectTypeHint(string content, out string typeName)
        {
            typeName = string.Empty;
            if (SymbolList.TypeHintToTypeName.Keys.Any(th => content.EndsWith(th)))
            {
                var lastChar = content.Substring(content.Length - 1);
                typeName = SymbolList.TypeHintToTypeName[lastChar];
                return true;
            }
            return false;
        }
    }


    public class UnreachableCaseInspectionRange : UnreachableCaseInspectionContext, IUnreachableCaseInspectionRange
    {
        private readonly bool _isValueRange;
        private readonly bool _isLTorGT;
        private readonly bool _isSingleValue;
        private readonly bool _isRelationalOp;
        private string _evalTypeName;
        private readonly IUnreachableCaseInspectionSummaryClauseFactory _summaryCoverageFactory;
        public UnreachableCaseInspectionRange(VBAParser.RangeClauseContext context, IUnreachableCaseInspectionValueResults inspValues, IUnreachableCaseInspectionSummaryClauseFactory factory, IUnreachableCaseInspectionValueFactory valueFactory) 
            : base(context, inspValues, factory, valueFactory)
        {
            _isValueRange = Context.HasChildToken(Tokens.To);
            _isLTorGT = Context.HasChildToken(Tokens.Is);
            _isRelationalOp = Context.TryGetChildContext<VBAParser.RelationalOpContext>(out _);
            _isSingleValue = !(_isValueRange || _isLTorGT || _isRelationalOp);
            _summaryCoverageFactory = factory;
        }

        public bool HasIncompatibleType { get; set; }
        public bool IsValueRange => _isValueRange;
        public bool IsLTorGT => _isLTorGT;
        public bool IsSingleValue => _isSingleValue;
        public bool IsRelationalOp => _isRelationalOp;
        public new ISummaryCoverage SummaryCoverage
        {
            get
            {
                if(_summaryCoverage is null || !_summaryCoverage.HasCoverage)
                {
                    _summaryCoverage = Coverage(_inspValues, EvaluationTypeName);
                }
                return _summaryCoverage;
            }
        }

        //NOTE: EvaluationTypeName used to make a SummaryCoverage object,
        //this may still be needed for SummaryCoverage<T> - 
        //it was removed for SummaryCoverage2<T>
        public string EvaluationTypeName { get; set; }
        public ParserRuleContext Result => FirstResultContext();
        public ParserRuleContext RangeStart => GetChild<VBAParser.SelectStartValueContext>();
        public ParserRuleContext RangeEnd => GetChild<VBAParser.SelectEndValueContext>();
        public string IsClauseSymbol => GetChild<VBAParser.ComparisonOperatorContext>().GetText();

        public ISummaryCoverage Coverage(IUnreachableCaseInspectionValueResults results, string evaluationTypeName)
        {
            _summaryCoverage = _summaryCoverageFactory.Create(evaluationTypeName, _factoryValue);
            try
            {
                if (IsSingleValue)
                {
                    _summaryCoverage.AddSingleValue(results.GetValue(Result));
                }
                else if (IsValueRange)
                {
                    _summaryCoverage.AddValueRange(results.GetValue(RangeStart), results.GetValue(RangeEnd));
                }
                else if (IsLTorGT)
                {
                    _summaryCoverage.AddIsClause(results.GetValue(Result), IsClauseSymbol);
                }
                else if (IsRelationalOp)
                {
                    _summaryCoverage.AddRelationalOp(results.GetValue(Result));
                }
            }
            catch (ArgumentException)
            {
                HasIncompatibleType = true;
            }
            return _summaryCoverage;
        }

        protected ParserRuleContext FirstResultContext()
        {
            var resultContexts = Context.children.Where(p => p is ParserRuleContext).Where(ch => IsResultContext((ParserRuleContext)ch)).Select(k => (ParserRuleContext)k);
            if (resultContexts.Any())
            {
                return resultContexts.First();
            }
            //TODO: exception?
            return null;
        }

        protected override bool IsResultContext<TContext>(TContext context)
        {
            return UnreachableCaseInspectionValueVisitor.IsMathContext(context)
                    || UnreachableCaseInspectionValueVisitor.IsLogicalContext(context)
                    || context is VBAParser.SelectStartValueContext
                    || context is VBAParser.SelectEndValueContext
                    || context is VBAParser.ParenthesizedExprContext
                    || context is VBAParser.SelectEndValueContext
                    || context is VBAParser.LExprContext
                    || context is VBAParser.LiteralExprContext;
        }
    }
}
