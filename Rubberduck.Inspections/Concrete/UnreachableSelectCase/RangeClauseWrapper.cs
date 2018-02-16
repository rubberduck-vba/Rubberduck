using Antlr4.Runtime;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Rubberduck.Inspections.Concrete
{
    //public class RangeClauseWrapper<T> where T : IComparable<T>
    //{
    //    private ContextValueResults<T> _ctxtValueResults;
    //    private readonly VBAParser.RangeClauseContext _rangeClause;
    //    private readonly bool _isValueRange;
    //    private readonly bool _isIsStatement;
    //    private readonly bool _isRelationalOpContext;
    //    private SummaryCoverage<T> _summaryCoverage;
    //    private IEnumerable<ParserRuleContext> _myValueResults;
    //    private IEnumerable<ParserRuleContext> _myUnresolvedResults;
    //    private bool _rangesHasMismatch;

    //    public RangeClauseWrapper(VBAParser.RangeClauseContext rangeClause, ContextValueResults<T> ctxtValueResults)
    //    {
    //        _ctxtValueResults = ctxtValueResults;
    //        _rangeClause = rangeClause;
    //        _rangesHasMismatch = false;

    //        _isValueRange = _rangeClause.HasChildToken(Tokens.To);

    //        _isIsStatement = (_rangeClause.HasChildToken(Tokens.Is) && _rangeClause.TryGetChildContext(out VBAParser.ComparisonOperatorContext _));

    //        _isRelationalOpContext = _rangeClause.TryGetChildContext(out VBAParser.RelationalOpContext _);

    //        _myValueResults = _rangeClause.GetDescendents<ParserRuleContext>().Where(ch => _ctxtValueResults.ValueResolvedContexts.Keys.Contains(ch));

    //        _myUnresolvedResults = _rangeClause.GetDescendents<ParserRuleContext>().Where(ch => _ctxtValueResults.VariableContexts.Keys.Contains(ch));

    //        _summaryCoverage = LoadSummaryCoverage(new SummaryCoverage<T>(_ctxtValueResults.Extents, _ctxtValueResults.TrueValue, _ctxtValueResults.FalseValue));
    //    }

    //    private bool IsValueRange => _isValueRange;
    //    private bool IsIsClause => _isIsStatement;
    //    private bool IsRelationalOp => _isRelationalOpContext;
    //    private bool IsSingleValues => !(IsValueRange || IsIsClause || IsRelationalOp);
    //    public bool CanBeInspected
    //    {
    //        get
    //        {
    //            return _summaryCoverage.HasCoverage || _summaryCoverage.HasExtents;
    //        }
    //    }

    //    public bool IsIncompatibleType
    //    {
    //        get
    //        {
    //            if (IsValueRange)
    //            {
    //                return _rangesHasMismatch;
    //            }
    //            var values = _myUnresolvedResults.Select(urc => _ctxtValueResults.VariableContexts[urc]);
    //            var result = !_myValueResults.Any() && values.All(v => !(v.DerivedTypeName.Equals(v.UseageTypeName)));
    //            return result;
    //        }
    //    }

    //    public SummaryCoverage<T> SummaryCoverage => _summaryCoverage;


    //    private SummaryCoverage<T> LoadSummaryCoverage(SummaryCoverage<T> summaryCoverage)
    //    {
    //        if (IsValueRange)
    //        {
    //            summaryCoverage = LoadValueRange(summaryCoverage);
    //        }
    //        else if (IsIsClause)
    //        {
    //            summaryCoverage = LoadIsClause(summaryCoverage);
    //        }
    //        else if (IsRelationalOp)
    //        {
    //            summaryCoverage = LoadRelationalOpClause(summaryCoverage);
    //        }
    //        else if (IsSingleValues)
    //        {
    //            summaryCoverage = LoadSingleValue(summaryCoverage);
    //        }
    //        return summaryCoverage;
    //    }

    //    private SummaryCoverage<T> LoadSingleValue(SummaryCoverage<T> summaryCoverage)
    //    {
    //        var ctxts = _rangeClause.GetChildren<ParserRuleContext>().Where(ch => _ctxtValueResults.ValueResolvedContexts.Keys.Contains(ch));

    //        if (ctxts.Any())
    //        {
    //            summaryCoverage.Add(_ctxtValueResults.ValueResolvedContexts[ctxts.First()]);
    //        }
    //        return summaryCoverage;
    //    }

    //    private SummaryCoverage<T> LoadValueRange(SummaryCoverage<T> summaryCoverage)
    //    {
    //        var startContext = _rangeClause.GetChild<VBAParser.SelectStartValueContext>();
    //        var endContext = _rangeClause.GetChild<VBAParser.SelectEndValueContext>();

    //        var hasStart = _ctxtValueResults.ValueResolvedContexts.TryGetValue(startContext, out T startVal);
    //        var hasEnd = _ctxtValueResults.ValueResolvedContexts.TryGetValue(endContext, out T endVal);

    //        if (hasStart && hasEnd)
    //        {
    //            summaryCoverage.AddRange(startVal, endVal);
    //        }

    //        if (!hasStart)
    //        {
    //            _rangesHasMismatch = RangeStartOrEndHasMismatch(startContext, _ctxtValueResults);
    //        }
    //        if (!hasEnd && !_rangesHasMismatch)
    //        {
    //            _rangesHasMismatch = RangeStartOrEndHasMismatch(endContext, _ctxtValueResults);
    //        }
    //        return summaryCoverage;
    //    }

    //    private static bool RangeStartOrEndHasMismatch(ParserRuleContext prCtxt, ContextValueResults<T> ctxtValues)
    //    {
    //        bool mismatchFound = false;
    //        foreach (var ctxt in prCtxt.GetChildren<ParserRuleContext>())
    //        {
    //            if (ctxtValues.VariableContexts.Keys.Contains(ctxt))
    //            {
    //                var value = ctxtValues.VariableContexts[ctxt];
    //                if (!mismatchFound)
    //                {
    //                    mismatchFound = !value.HasValueAs(ctxtValues.EvaluationTypeName);
    //                }
    //            }
    //        }
    //        return mismatchFound;
    //    }

    //    //RelationalOp statements like x < 100, 100 <> x, where x is a constant value
    //    private SummaryCoverage<T> LoadRelationalOpClause(SummaryCoverage<T> summaryCoverage)
    //    {
    //        if (_rangeClause.TryGetChildContext(out VBAParser.RelationalOpContext relOpCtxt))
    //        {
    //            if (_ctxtValueResults.ValueResolvedContexts.Keys.Contains(relOpCtxt))
    //            {
    //                summaryCoverage.RelationalOps.Add(_ctxtValueResults.ValueResolvedContexts[relOpCtxt]);
    //            }
    //            else
    //            {
    //                summaryCoverage.RelationalOps.Add(relOpCtxt.GetText());
    //            }
    //        }
    //        return summaryCoverage;
    //    }

    //    private SummaryCoverage<T> LoadIsClause(SummaryCoverage<T> summaryCoverage)
    //    {
    //        var ctxts = _rangeClause.GetChildren<ParserRuleContext>().Where(ch => _ctxtValueResults.ValueResolvedContexts.Keys.Contains(ch));
    //        if (_rangeClause.HasChildToken(Tokens.Is))
    //        {
    //            var compOpContext = _rangeClause.GetChild<VBAParser.ComparisonOperatorContext>();
    //            summaryCoverage.AddIsClauseResult(compOpContext.GetText(), _ctxtValueResults.ValueResolvedContexts[(ParserRuleContext)ctxts.First()]);
    //            return summaryCoverage;
    //        }
    //        return summaryCoverage;
    //    }
    //}
}
