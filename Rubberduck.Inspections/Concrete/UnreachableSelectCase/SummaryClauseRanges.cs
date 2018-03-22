using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Rubberduck.Inspections.Concrete.UnreachableSelectCase
{
    public class SummaryClauseRanges<T> : SummaryClauseBase<T> where T : IComparable<T>
    {
        private List<Tuple<T, T>> _ranges;

        public SummaryClauseRanges(Func<IUnreachableCaseInspectionValue, T> tConverter) : base(tConverter)
        {
            RangeClauses = new List<ISummaryClauseRange<T>>();
            _ranges = new List<Tuple<T, T>>();
        }

        public List<ISummaryClauseRange<T>> RangeClauses { set; get; }
        public override bool HasCoverage => Any();
        private bool Any() => RangeClauses.Any() || _ranges.Any();

        public void Add(ISummaryClauseRange<T> candidate)
        {
            if (Covers(candidate))
            {
                return;
            }
            AddRange(candidate);
        }

        public bool Covers(ISummaryClauseRange<T> candidate)
        {
            if (!HasCoverage)
            {
                return false;
            }
            return RangeClauses.Any(rg => rg.Covers(candidate.Start) && rg.Covers(candidate.End));
        }

        public override bool Covers(T candidate)
        {
            if (!HasCoverage)
            {
                return false;
            }
            return RangeClauses.Any(rg => rg.Start.CompareTo(candidate) <= 0 && rg.End.CompareTo(candidate) >= 0)
                || _ranges.Any(rg => rg.Item1.CompareTo(candidate) <= 0 && rg.Item2.CompareTo(candidate) >= 0);
        }

        public void RemoveIfCoveredBy(SummaryClauseRanges<T> ranges)
        {
            var toRemove = new List<ISummaryClauseRange<T>>();
            for (var idx = 0; idx < RangeClauses.Count; idx++)
            {
                var rangeClause = RangeClauses[idx];
                if (ranges.RangeClauses.Any(rg => rg.Covers(rangeClause)))
                {
                    toRemove.Add(rangeClause);
                }
            }
            Remove(toRemove);
        }

        public void Remove(List<ISummaryClauseRange<T>> rangesToRemove)
        {
            foreach( var range in rangesToRemove)
            {
                Remove(range);
            }
        }

        public void Remove(ISummaryClauseRange<T> rangeToRemove)
        {
            RangeClauses.Remove(rangeToRemove);
        }

        public override string ToString()
        {
            if (!RangeClauses.Any())
            {
                return string.Empty;
            }
            const string prefix = "Range=";
            var result = string.Empty;
            foreach (var range in RangeClauses)
            {
                result = result.Length > 0 ? $"{result},{range.ToString()}" : $"{range.ToString()}";
            }
            return $"{prefix}{result}";
        }

        public List<long> AsIntegerNumbers
        {
            get
            {
                var results = new List<long>();
                if (ContainsIntegerNumbers)
                {
                    foreach( var range in RangeClauses)
                    {
                        results.AddRange(range.AsIntegerNumbers);
                    }
                }
                return results;
            }
        }

        private void AddRange(ISummaryClauseRange<T> rangeClause)
        {
            if (!RangeClauses.Any())
            {
                RangeClauses.Add(rangeClause);
                return;
            }

            if (Covers(rangeClause))
            {
                return;
            }

            RangeClauses.Add(rangeClause);

            var somethingRemovedOrCombined = false;
            do
            {
                somethingRemovedOrCombined = false;
                var indexesToRemove = CombineAdjacentRanges();
                if (indexesToRemove.Any())
                {
                    foreach (var idx in indexesToRemove)
                    {
                        somethingRemovedOrCombined = true;
                        RangeClauses.RemoveAt(idx);
                    }
                }

                indexesToRemove = CombineOverlappingRanges();
                if (indexesToRemove.Any())
                {
                    foreach (var idx in indexesToRemove)
                    {
                        somethingRemovedOrCombined = true;
                        RangeClauses.RemoveAt(idx);
                    }
                }
            } while (somethingRemovedOrCombined);
        }

        private List<int> CombineAdjacentRanges()
        {
            var indexesToRemove = new List<int>();
            if (ContainsIntegerNumbers)
            {
                var nextIndex = 1;
                for (var idx = 0; nextIndex < RangeClauses.Count; idx++, nextIndex = idx + 1)
                {
                    if (RangeClauses[nextIndex].IsAdjacent(RangeClauses[idx]))
                    {
                        RangeClauses[nextIndex].AppendRange(RangeClauses[idx]);
                        indexesToRemove.Add(idx);
                    }
                }
            }
            return indexesToRemove;
        }

        private List<int> CombineOverlappingRanges()
        {
            var indexesToRemove = new List<int>();
            var nextIndex = 1;
            for (var idx = 0; nextIndex < RangeClauses.Count; idx++, nextIndex = idx + 1)
            {
                if (RangeClauses[nextIndex].Overlaps(RangeClauses[idx]))
                {
                    RangeClauses[nextIndex].RemoveOverlap(RangeClauses[idx]);
                    indexesToRemove.Add(idx);
                }
            }
            return indexesToRemove;
        }

    }
}
