using System;
using System.Collections.Generic;
using System.Linq;

namespace Rubberduck.Inspections.Concrete.UnreachableCaseInspection
{
    public class ExpressionFilterIntegral : ExpressionFilter<long>
    {
        public ExpressionFilterIntegral(StringToValueConversion<long> converter) : base(converter) { }

        protected override bool AddValueRange(RangeValues<long> range)
        {
            var addsRange = base.AddValueRange(range);

            ConcatenateExistingRanges();
            RemoveRangesCoveredByLimits();
            RemoveSingleValuesCoveredByRanges();
            return addsRange;
        }

        public override bool FiltersAllValues => base.FiltersAllValues
            || Limits.HasMinAndMaxLimits && (Limits.Maximum - Limits.Minimum + 1 <= RangesValuesCount + SingleValues.Count());

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
            var concatenatedRanges = new List<RangeValues<long>>();
            var indexesToRemove = new List<int>();
            var sortedRanges = Ranges.Select(rg => new RangeValues<long>(rg.Start, rg.End)).OrderBy(k => k.Start).ToList();
            for (int idx = sortedRanges.Count() - 1; idx > 0;)
            {
                if (sortedRanges[idx].Start == sortedRanges[idx - 1].End || sortedRanges[idx].Start - sortedRanges[idx - 1].End == 1)
                {
                    concatenatedRanges.Add(new RangeValues<long>(sortedRanges[idx - 1].Start, sortedRanges[idx].End));
                    indexesToRemove.Add(idx);
                    indexesToRemove.Add(idx - 1);
                    idx = -1;
                }
                idx--;
            }
            //rebuild _ranges retaining the original order except placing the concatenated
            //range added to the end
            if (concatenatedRanges.Any())
            {
                int idx = 0;
                var allRanges = new Dictionary<int, RangeValues<long>>();
                foreach( var range in Ranges)
                {
                    allRanges.Add(idx++, range);
                }

                indexesToRemove.ForEach(id => sortedRanges.RemoveAt(id));

                var tRanges = new List<RangeValues<long>>();
                sortedRanges.ForEach(sr => tRanges.Add(new RangeValues<long>(sr.Start,sr.End)));

                concatenatedRanges.ForEach(sr => tRanges.Add(new RangeValues<long>(sr.Start, sr.End)));

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
