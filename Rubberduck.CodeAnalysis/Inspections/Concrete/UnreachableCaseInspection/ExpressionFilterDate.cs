using Rubberduck.Parsing.Grammar;
using System;
using System.Linq;

namespace Rubberduck.Inspections.Concrete.UnreachableCaseInspection
{
    public class ExpressionFilterDate : ExpressionFilter<ComparableDateValue>
    {
        public ExpressionFilterDate(TokenToValue<ComparableDateValue> converter) : base(converter, Tokens.Date) { }

        public override bool FiltersAllValues => base.FiltersAllValues
            || Limits.HasMinAndMaxLimits && (Limits.Maximum.AsDecimal - Limits.Minimum.AsDecimal + 1 <= RangesValuesCount + SingleValues.Count());

        private long RangesValuesCount => Ranges.Sum(rg => Convert.ToInt64(rg.End.AsDecimal) - Convert.ToInt64(rg.Start.AsDecimal) + 1);
    }
}
