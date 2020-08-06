using System;
using System.Collections.Generic;
using System.Linq;

namespace Rubberduck.CodeAnalysis.Inspections.Concrete.UnreachableCaseEvaluation
{
    internal class ExpressionFilterIntegral : ExpressionFilter<long>
    {
        public ExpressionFilterIntegral(string valueType, Func<string,long> parser) 
            : base(valueType, parser) { }

        public override bool FiltersAllValues => base.FiltersAllValues
            || Limits.HasMinAndMaxLimits && (Limits.Maximum - Limits.Minimum + 1 <= RangesValuesCount + SingleValues.Count());

        protected override bool AddValueRange(RangeOfValues rov)
        {
            var addsRange = base.AddValueRange(rov);

            ConcatenateExistingRanges();
            Ranges.RemoveWhere(rg => Limits.FiltersRange(rg.Start, rg.End));
            SingleValues.RemoveWhere(sv => Ranges.Any(rg => rg.Filters(sv)));
            return addsRange;
        }

        private long RangesValuesCount => Ranges.Sum(rg => Convert.ToInt64(rg.End) - Convert.ToInt64(rg.Start) + 1);

        private void ConcatenateExistingRanges()
        {
            if (Ranges.Count() > 1)
            {
                int preConcatentateCount;
                do
                {
                    preConcatentateCount = Ranges.Count();
                    ConcatenateRanges();
                } while (Ranges.Count() < preConcatentateCount && Ranges.Count() > 1);
            }
        }

        private void ConcatenateRanges()
        {
            var concatenatedRanges = new List<RangeOfValues>();
            var indexesToRemove = new List<int>();
            var sortedRanges = Ranges.OrderBy(k => k.Start).ToList();
            for (int idx = sortedRanges.Count() - 1; idx > 0;)
            {
                if (sortedRanges[idx].Start == sortedRanges[idx - 1].End || sortedRanges[idx].Start - sortedRanges[idx - 1].End == 1)
                {
                    concatenatedRanges.Add( new RangeOfValues(sortedRanges[idx - 1].Start, sortedRanges[idx].End));
                    indexesToRemove.Add(idx);
                    indexesToRemove.Add(idx - 1);
                    idx = -1;
                }
                idx--;
            }
            //rebuild _ranges retaining the original order placing the concatenated ranges at the end
            if (concatenatedRanges.Any())
            {
                int idx = 0;
                var allRanges = new Dictionary<int, RangeOfValues>();
                foreach ( var range in Ranges)
                {
                    allRanges.Add(idx++, range);
                }

                indexesToRemove.ForEach(id => sortedRanges.RemoveAt(id));

                var tRanges = new List<RangeOfValues>();
                sortedRanges.ForEach(sr => tRanges.Add(sr));

                concatenatedRanges.ForEach(sr => tRanges.Add(new RangeOfValues(sr.Start, sr.End)));

                Ranges.Clear();
                var removalKeys = allRanges.Keys.Where(k => tRanges.Contains(allRanges[k]));
                foreach(var nk in removalKeys)
                {
                    Ranges.Add(allRanges[nk]);
                    tRanges.Remove(allRanges[nk]);
                }

                tRanges.ForEach(tr => Ranges.Add(tr));
            }
        }
    }
}
