using Rubberduck.Parsing.Grammar;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Rubberduck.Inspections.Concrete.UnreachableCaseInspection
{
    public class ExpressionFilterDate : ExpressionFilter<DateValueIComparableDecorator>
    {
        public ExpressionFilterDate(StringToValueConversion<DateValueIComparableDecorator> converter) : base(converter, Tokens.Date) { }

        public override bool FiltersAllValues => base.FiltersAllValues
            || Limits.HasMinAndMaxLimits && (Limits.Maximum.AsDecimal - Limits.Minimum.AsDecimal + 1 <= RangesValuesCount + SingleValues.Count());

        private long RangesValuesCount => Ranges.Sum(rg => Convert.ToInt64(rg.End.AsDecimal) - Convert.ToInt64(rg.Start.AsDecimal) + 1);
    }
}
