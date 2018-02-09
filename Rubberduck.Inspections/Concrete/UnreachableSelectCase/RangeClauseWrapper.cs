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
    public class RangeClauseWrapper<T> where T : IComparable<T>
    {
        private ContextValueResults<T> _ctxtValueResults;
        private readonly VBAParser.RangeClauseContext _rangeClause;
        private readonly bool _isValueRange;
        private readonly bool _isIsStatement;
        private Dictionary<ParserRuleContext, string> _relationalOpContexts;
        private SummaryCoverage<T> _summaryCoverage;
        private IEnumerable<ParserRuleContext> _myValueResults;
        private IEnumerable<ParserRuleContext> _myUnresolvedResults;
        private bool _rangeClauseHasMismatch;

        public RangeClauseWrapper(VBAParser.RangeClauseContext rangeClause, ContextValueResults<T> ctxtValueResults)
        {
            _ctxtValueResults = ctxtValueResults;
            _rangeClause = rangeClause;
            _rangeClauseHasMismatch = false;

            _isValueRange = _rangeClause.HasChildToken(Tokens.To);

            _isIsStatement = (_rangeClause.HasChildToken(Tokens.Is) && _rangeClause.TryGetChildContext(out VBAParser.ComparisonOperatorContext _))
                    || _rangeClause.TryGetChildContext(out VBAParser.RelationalOpContext _);

            _relationalOpContexts = new Dictionary<ParserRuleContext, string>();

            _myValueResults = _rangeClause.GetDescendents<ParserRuleContext>().Where(ch => _ctxtValueResults.ValueResolvedContexts.Keys.Contains(ch));

            _myUnresolvedResults = _rangeClause.GetDescendents<ParserRuleContext>().Where(ch => _ctxtValueResults.UnresolvedContexts.Keys.Contains(ch));

            _summaryCoverage = LoadSummaryCoverage(new SummaryCoverage<T>(_ctxtValueResults.Extents));
        }

        private bool IsValueRangeRangeClause => _isValueRange;
        private bool IsIsRangeClause => _isIsStatement;
        private bool IsSingleValueRangeClause => !(IsValueRangeRangeClause || IsIsRangeClause);
        public bool CanBeInspected
        {
            get
            {
                var relOpsOK = _relationalOpContexts.Keys.Any() ? _relationalOpContexts.Keys.Count < 2 : true;
                return _myValueResults.Any() || relOpsOK;
            }
        }

        public bool IsMismatch
        {
            get
            {
                if (IsValueRangeRangeClause)
                {
                    return _rangeClauseHasMismatch;
                }
                var values = _myUnresolvedResults.Select(urc => _ctxtValueResults.UnresolvedContexts[urc]);
                var result = !_myValueResults.Any() && values.All(v => !(v.DerivedTypeName.Equals(v.UseageTypeName)));
                return result;
            }
        }

        public SummaryCoverage<T> SummaryCoverage => _summaryCoverage;


        private SummaryCoverage<T> LoadSummaryCoverage(SummaryCoverage<T> summaryCoverage)
        {
            if (IsValueRangeRangeClause)
            {
                summaryCoverage = LoadValueRange(summaryCoverage);
            }
            else if (IsIsRangeClause)
            {
                summaryCoverage = LoadIsClause(summaryCoverage);
            }
            else if (IsSingleValueRangeClause)
            {
                summaryCoverage = LoadSingleValue(summaryCoverage);
            }
            return summaryCoverage;
        }

        private SummaryCoverage<T> LoadSingleValue(SummaryCoverage<T> summaryCoverage)
        {
            var ctxts = _rangeClause.GetChildren<ParserRuleContext>().Where(ch => _ctxtValueResults.ValueResolvedContexts.Keys.Contains(ch));

            if (ctxts.Any())
            {
                summaryCoverage.Add(_ctxtValueResults.ValueResolvedContexts[ctxts.First()]);
            }
            return summaryCoverage;
        }

        private SummaryCoverage<T> LoadValueRange(SummaryCoverage<T> summaryCoverage)
        {
            var startContext = _rangeClause.GetChild<VBAParser.SelectStartValueContext>();
            var endContext = _rangeClause.GetChild<VBAParser.SelectEndValueContext>();

            var hasStart = _ctxtValueResults.ValueResolvedContexts.TryGetValue(startContext, out T startVal);
            var hasEnd = _ctxtValueResults.ValueResolvedContexts.TryGetValue(endContext, out T endVal);

            if (hasStart && hasEnd)
            {
                summaryCoverage.AddRange(startVal, endVal);
            }

            if (!hasStart)
            {
                _rangeClauseHasMismatch = RangeStartOrEndHasMismatch(startContext, _ctxtValueResults);
            }
            if (!hasEnd && !_rangeClauseHasMismatch)
            {
                _rangeClauseHasMismatch = RangeStartOrEndHasMismatch(endContext, _ctxtValueResults);
            }
            return summaryCoverage;
        }

        private static bool RangeStartOrEndHasMismatch(ParserRuleContext prCtxt, ContextValueResults<T> ctxtValues)
        {
            bool mismatchFound = false;
            foreach (var ctxt in prCtxt.GetChildren<ParserRuleContext>())
            {
                if (ctxtValues.UnresolvedContexts.Keys.Contains(ctxt))
                {
                    var value = ctxtValues.UnresolvedContexts[ctxt];
                    if (!mismatchFound)
                    {
                        mismatchFound = !value.HasValueAs(ctxtValues.EvaluationTypeName);
                    }
                }
            }
            return mismatchFound;
        }

        private SummaryCoverage<T> LoadIsClause(SummaryCoverage<T> summaryCoverage)
        {
            var ctxts = _rangeClause.GetChildren<ParserRuleContext>().Where(ch => _ctxtValueResults.ValueResolvedContexts.Keys.Contains(ch));
            if (ctxts.Count() == 1 && _rangeClause.HasChildToken(Tokens.Is))
            {
                //Is Statements
                var compOpContext = _rangeClause.GetChild<VBAParser.ComparisonOperatorContext>();
                summaryCoverage.AddIsClauseResult(compOpContext.GetText(), _ctxtValueResults.ValueResolvedContexts[(ParserRuleContext)ctxts.First()]);
                return summaryCoverage;
            }
            //RelationalOp statements like x < 100, 100 <> x
            else if (_rangeClause.TryGetChildContext(out VBAParser.RelationalOpContext relOpCtxt))
            {
                if (!_ctxtValueResults.ValueResolvedContexts.Keys.Contains(relOpCtxt))
                {
                    var relOpContexts = relOpCtxt.children.Where(ch => !(ch is VBAParser.WhiteSpaceContext)).ToList();
                    for (var idx = 0; idx < relOpContexts.Count(); idx++)
                    {
                        var ctxt = relOpContexts[idx];
                        if (_ctxtValueResults.ValueResolvedContexts.Keys.Contains(ctxt))
                        {
                            var value = _ctxtValueResults.ValueResolvedContexts[(ParserRuleContext)ctxt];
                            var opSymbol = relOpCtxt.children.Where(ch => SummaryCoverage<T>.BinaryLogicalOps.Keys.Contains(ch.GetText())).First().GetText();
                            if (idx == 0)
                            {
                                //100 < x: when the value is the first child, the expression's opSymbol
                                //needs to be converted to represent x < 100
                                summaryCoverage.AddIsClauseResult(SummaryCoverage<T>.AlgebraicLogicalInversions[opSymbol], value);
                            }
                            else
                            {
                                summaryCoverage.AddIsClauseResult(opSymbol, value);
                            }
                        }
                        else if (ctxt is ParserRuleContext prCtxt)
                        {
                            _relationalOpContexts.Add(prCtxt, prCtxt.GetText());
                        }
                    }
                }
            }
            return summaryCoverage;
        }
    }
}
