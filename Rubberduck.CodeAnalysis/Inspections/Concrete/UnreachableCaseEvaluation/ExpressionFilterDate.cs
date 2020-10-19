using System;
using System.Linq;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Refactoring.ParseTreeValue;

namespace Rubberduck.CodeAnalysis.Inspections.Concrete.UnreachableCaseEvaluation
{
    internal class ExpressionFilterDate : ExpressionFilter<ComparableDateValue>
    {
        public ExpressionFilterDate() : base(Tokens.Date, ComparableDateValue.Parse) { }

        public override bool FiltersAllValues => base.FiltersAllValues
            || Limits.HasMinAndMaxLimits && (Limits.Maximum.AsDecimal - Limits.Minimum.AsDecimal + 1 <= RangesValuesCount + SingleValues.Count());

        private long RangesValuesCount => Ranges.Sum(rg => Convert.ToInt64(rg.End.AsDecimal) - Convert.ToInt64(rg.Start.AsDecimal) + 1);
    }
}
