using System;
using System.Linq;
using Antlr4.Runtime;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;

namespace Rubberduck.Inspections.Concrete.UnreachableCaseInspection
{
    public interface IUnreachableCaseInspectionRange
    {
        string EvaluationTypeName { get; }
        bool HasIncompatibleType { set; get; }
        bool HasCoverage { get; }
        bool IsValueRange { get; }
        bool IsLTorGT { get; }
        bool IsSingleValue { get; }
        bool IsRelationalOp { get; }
        ParserRuleContext Result { get; }
        ParserRuleContext RangeStart { get; }
        ParserRuleContext RangeEnd { get; }
        string IsClauseSymbol { get; }
        bool IsReachable(IUCIRangeClauseFilter filter);
        IUCIRangeClauseFilter Coverage(IUCIValueResults results, string evaluationTypeName);
        IUCIRangeClauseFilter RangeClause { set; get; }
    }

    public class UnreachableCaseInspectionRange : UnreachableCaseInspectionContext, IUnreachableCaseInspectionRange
    {
        private readonly bool _isValueRange;
        private readonly bool _isLTorGT;
        private readonly bool _isSingleValue;
        private readonly bool _isRelationalOp;
        private readonly IUCIRangeClauseFilterFactory _rangeFilterFactory;
        private string _evalTypeName;

        public UnreachableCaseInspectionRange(VBAParser.RangeClauseContext context, IUCIValueResults inspValues, IUCIRangeClauseFilterFactory factory, IUCIValueFactory valueFactory) 
            : base(context, inspValues, factory, valueFactory)
        {
            _isValueRange = Context.HasChildToken(Tokens.To);
            _isLTorGT = Context.HasChildToken(Tokens.Is);
            _isRelationalOp = Context.TryGetChildContext<VBAParser.RelationalOpContext>(out _);
            _isSingleValue = !(_isValueRange || _isLTorGT || _isRelationalOp);
            _rangeFilterFactory = factory;
            _evalTypeName = string.Empty;
            RangeClause = _rangeFilterFactory.Create(Tokens.Long, _factoryValue, _rangeFilterFactory);
        }

        public string EvaluationTypeName
        {
            set
            {
                if (value != _evalTypeName)
                {
                    _evalTypeName = value;
                    RangeClause = Coverage();
                }
            }
            get => _evalTypeName;
        }

        public bool HasIncompatibleType { get; set; }
        public bool HasCoverage => RangeClause.HasCoverage;
        public bool IsValueRange => _isValueRange;
        public bool IsLTorGT => _isLTorGT;
        public bool IsSingleValue => _isSingleValue;
        public bool IsRelationalOp => _isRelationalOp;
        public IUCIRangeClauseFilter RangeClause { set; get; }
        //public ISummaryCoverage Filter(ISummaryCoverage filter);
        public IUCIValueResults InspectionResults { set; get; }
        public ParserRuleContext Result => FirstResultContext();
        public ParserRuleContext RangeStart => GetChild<VBAParser.SelectStartValueContext>();
        public ParserRuleContext RangeEnd => GetChild<VBAParser.SelectEndValueContext>();
        public string IsClauseSymbol => GetChild<VBAParser.ComparisonOperatorContext>().GetText();

        public bool IsReachable(IUCIRangeClauseFilter filter)
        {
            var inspectedCoverage = Coverage();
            RangeClause = inspectedCoverage.FilterUnreachableClauses(filter);
            return RangeClause.HasCoverage;
        }

        private IUCIRangeClauseFilter Coverage()
        {
            var rangeClauseCoverage = _rangeFilterFactory.Create(EvaluationTypeName, _factoryValue, _rangeFilterFactory);
            try
            {
                if (IsSingleValue)
                {
                    rangeClauseCoverage.AddSingleValue(_inspValues.GetValue(Result));
                }
                else if (IsValueRange)
                {
                    rangeClauseCoverage.AddValueRange(_inspValues.GetValue(RangeStart), _inspValues.GetValue(RangeEnd));
                }
                else if (IsLTorGT)
                {
                    rangeClauseCoverage.AddIsClause(_inspValues.GetValue(Result), IsClauseSymbol);
                }
                else if (IsRelationalOp)
                {
                    rangeClauseCoverage.AddRelationalOp(_inspValues.GetValue(Result));
                }
            }
            catch (ArgumentException)
            {
                HasIncompatibleType = true;
            }
            return rangeClauseCoverage;
        }

        public IUCIRangeClauseFilter Coverage(IUCIValueResults results, string evaluationTypeName)
        {
            RangeClause = _rangeFilterFactory.Create(evaluationTypeName, _factoryValue, _rangeFilterFactory);
            try
            {
                if (IsSingleValue)
                {
                    RangeClause.AddSingleValue(results.GetValue(Result));
                }
                else if (IsValueRange)
                {
                    RangeClause.AddValueRange(results.GetValue(RangeStart), results.GetValue(RangeEnd));
                }
                else if (IsLTorGT)
                {
                    RangeClause.AddIsClause(results.GetValue(Result), IsClauseSymbol);
                }
                else if (IsRelationalOp)
                {
                    RangeClause.AddRelationalOp(results.GetValue(Result));
                }
            }
            catch (ArgumentException)
            {
                HasIncompatibleType = true;
            }
            return RangeClause;
        }

        public override string ToString()
        {
            return $"{Context.GetText()} ({RangeClause.ToString()})";
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
            return UCIParseTreeValueVisitor.IsMathContext(context)
                    || UCIParseTreeValueVisitor.IsLogicalContext(context)
                    || context is VBAParser.SelectStartValueContext
                    || context is VBAParser.SelectEndValueContext
                    || context is VBAParser.ParenthesizedExprContext
                    || context is VBAParser.SelectEndValueContext
                    || context is VBAParser.LExprContext
                    || context is VBAParser.LiteralExprContext;
        }
    }
}
