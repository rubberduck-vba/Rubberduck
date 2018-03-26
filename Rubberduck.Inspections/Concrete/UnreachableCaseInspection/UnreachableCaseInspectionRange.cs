using System;
using System.Collections.Generic;
using System.Linq;
using Antlr4.Runtime;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;

namespace Rubberduck.Inspections.Concrete.UnreachableCaseInspection
{
    public interface IUnreachableCaseInspectionRange
    {
        string EvaluationTypeName { set;  get; }
        bool HasIncompatibleType { set; get; }
        bool IsUnreachable { set; get; }
        bool IsReachable(IUCIRangeClauseFilter filter);
        IUCIRangeClauseFilter AsFilter { set; get; }
        List<ParserRuleContext> ResultContexts { get; }
    }

    public class UnreachableCaseInspectionRange : UnreachableCaseInspectionContext, IUnreachableCaseInspectionRange
    {
        private readonly bool _isValueRange;
        private readonly bool _isLTorGT;
        private readonly bool _isSingleValue;
        private readonly bool _isRelationalOp;
        private string _evalTypeName;

        public UnreachableCaseInspectionRange(VBAParser.RangeClauseContext context, IUCIValueResults inspValues, IUnreachableCaseInspectionFactoryFactory factoryFactory)
            : base(context, inspValues, factoryFactory)
        {
            _isValueRange = Context.HasChildToken(Tokens.To);
            _isLTorGT = Context.HasChildToken(Tokens.Is);
            _isRelationalOp = Context.TryGetChildContext<VBAParser.RelationalOpContext>(out _);
            _isSingleValue = !(_isValueRange || _isLTorGT || _isRelationalOp);
            _evalTypeName = string.Empty;
            IsUnreachable = false;
            AsFilter = FilterFactory.Create(Tokens.Long, ValueFactory);
        }

        public string EvaluationTypeName
        {
            set
            {
                if (value != _evalTypeName)
                {
                    _evalTypeName = value;
                    AsFilter = AddFilterContent();
                }
            }
            get => _evalTypeName;
        }

        public bool HasIncompatibleType { get; set; }

        public bool IsUnreachable { set; get; }

        public IUCIRangeClauseFilter AsFilter { set; get; }

        public List<ParserRuleContext> ResultContexts
        {
            get
            {
                var results = new List<ParserRuleContext>();
                if(!TryGetFirstResultContext(out ParserRuleContext resultContext))
                {
                    return results;
                }

                if (_isValueRange)
                {
                    results.Add(SelectStartValue);
                    results.Add(SelectEndValue);
                }
                else
                {
                    results.Add(resultContext);
                }
                return results;
            }
        }

        public bool IsReachable(IUCIRangeClauseFilter filter)
        {
            if (filter.FiltersAllValues)
            {
                IsUnreachable = true;
                return false;
            }

            var inspectedCoverage = AddFilterContent();
            AsFilter = inspectedCoverage.FilterUnreachableClauses(filter);
            IsUnreachable = !AsFilter.HasCoverage;
            return AsFilter.HasCoverage;
        }

        public override string ToString()
        {
            return $"{Context.GetText()} ({AsFilter.ToString()})";
        }

        private bool IsResultContext<TContext>(TContext context)
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

        private string IsClauseSymbol => Context.GetChild<VBAParser.ComparisonOperatorContext>().GetText(); // GetChild<VBAParser.ComparisonOperatorContext>().GetText();

        private ParserRuleContext SelectStartValue => Context.GetChild<VBAParser.SelectStartValueContext>();

        private ParserRuleContext SelectEndValue => Context.GetChild<VBAParser.SelectEndValueContext>();

        private IUCIRangeClauseFilter AddFilterContent()
        {
            var rangeClauseFilter = FilterFactory.Create(EvaluationTypeName, ValueFactory);
            try
            {
                if (!TryGetFirstResultContext(out ParserRuleContext resultContext))
                {
                    return rangeClauseFilter;
                 }

                if (_isSingleValue)
                {
                    rangeClauseFilter.AddSingleValue(ParseTreeValueResults.GetValue(resultContext));
                }
                else if (_isValueRange)
                {
                    rangeClauseFilter.AddValueRange(ParseTreeValueResults.GetValue(SelectStartValue), ParseTreeValueResults.GetValue(SelectEndValue));
                }
                else if (_isLTorGT)
                {
                    rangeClauseFilter.AddIsClause(ParseTreeValueResults.GetValue(resultContext), IsClauseSymbol);
                }
                else if (_isRelationalOp)
                {
                    rangeClauseFilter.AddRelationalOp(ParseTreeValueResults.GetValue(resultContext));
                }
            }
            catch (ArgumentException)
            {
                HasIncompatibleType = true;
            }
            return rangeClauseFilter;
        }

        private bool TryGetFirstResultContext( out ParserRuleContext result)
        {
            result = null;
            var resultContexts = Context.children.Where(p => p is ParserRuleContext).Where(ch => IsResultContext((ParserRuleContext)ch)).Select(k => (ParserRuleContext)k);
            if (resultContexts.Any())
            {
                result = resultContexts.First();
                return true;
            }
            return false;
        }
    }
}
