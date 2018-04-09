using System;
using System.Collections.Generic;
using System.Linq;
using Antlr4.Runtime;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;

namespace Rubberduck.Inspections.Concrete.UnreachableCaseInspection
{
    public interface IRangeClauseContextWrapper
    {
        string EvaluationTypeName { set;  get; }
        bool HasIncompatibleType { set; get; }
        bool IsUnreachable { set; get; }
        bool IsReachable(IRangeClauseFilter filter);
        IRangeClauseFilter AsFilter { set; get; }
        List<ParserRuleContext> ResultContexts { get; }
    }

    public class RangeClauseContextWrapper : ContextWrapperBase, IRangeClauseContextWrapper
    {
        private readonly bool _isValueRange;
        private readonly bool _isLTorGT;
        private readonly bool _isSingleValue;
        private readonly bool _isRelationalOp;
        private string _evalTypeName;

        public RangeClauseContextWrapper(VBAParser.RangeClauseContext context, IParseTreeVisitorResults inspValues, IUnreachableCaseInspectionFactoryProvider factoryFactory)
            : base(context, inspValues, factoryFactory)
        {
            _isValueRange = !(context.TO() is null);
            _isLTorGT = !(context.IS() is null);
            _isRelationalOp = Context.children.Any(ch => ch is ParserRuleContext && ParseTreeValueVisitor.IsLogicalContext(ch));
            _isSingleValue = !(_isValueRange || _isLTorGT || _isRelationalOp);
            _evalTypeName = string.Empty;
            IsUnreachable = false;
            AsFilter = FilterFactory.Create(Tokens.Long, ValueFactory);
        }

        public RangeClauseContextWrapper(VBAParser.RangeClauseContext context, string evalTypeName, IParseTreeVisitorResults inspValues, IUnreachableCaseInspectionFactoryProvider factoryFactory)
            : base(context, inspValues, factoryFactory)
        {
            _isValueRange = !(context.TO() is null);
            _isLTorGT = !(context.IS() is null);
            _isRelationalOp = Context.children.Any(ch => ch is ParserRuleContext && ParseTreeValueVisitor.IsLogicalContext(ch));

            _isSingleValue = !(_isValueRange || _isLTorGT || _isRelationalOp);
            IsUnreachable = false;
            AsFilter = FilterFactory.Create(evalTypeName, ValueFactory);
            _evalTypeName = evalTypeName;
            AsFilter = AddFilterContent();
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

        public IRangeClauseFilter AsFilter { set; get; }

        public List<ParserRuleContext> ResultContexts
        {
            get
            {
                var results = new List<ParserRuleContext>();
                if (!TryGetFirstResultContext(out ParserRuleContext resultContext))
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

        public bool IsReachable(IRangeClauseFilter filter)
        {
            if (filter.FiltersAllValues)
            {
                IsUnreachable = true;
                return false;
            }

            var inspectedCoverage = AddFilterContent();
            AsFilter = inspectedCoverage.FilterUnreachableClauses(filter);
            IsUnreachable = !AsFilter.ContainsFilters;
            return AsFilter.ContainsFilters;
        }

        public override string ToString()
        {
            return $"{Context.GetText()} ({AsFilter.ToString()})";
        }

        private bool IsResultContext<TContext>(TContext context)
        {
            return ParseTreeValueVisitor.IsMathContext(context)
                    || ParseTreeValueVisitor.IsLogicalContext(context)
                    || context is VBAParser.SelectStartValueContext
                    || context is VBAParser.SelectEndValueContext
                    || context is VBAParser.ParenthesizedExprContext
                    || context is VBAParser.SelectEndValueContext
                    || context is VBAParser.LExprContext
                    || context is VBAParser.LiteralExprContext;
        }

        private string IsClauseSymbol => Context.GetChild<VBAParser.ComparisonOperatorContext>().GetText();

        private ParserRuleContext SelectStartValue => Context.GetChild<VBAParser.SelectStartValueContext>();

        private ParserRuleContext SelectEndValue => Context.GetChild<VBAParser.SelectEndValueContext>();

        private IRangeClauseFilter AddFilterContent()
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
