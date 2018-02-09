using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Rubberduck.Inspections.Concrete
{
    public class SummaryClauseRanges<T> : SummaryClauseBase<T> where T : IComparable<T>
    {
        public SummaryClauseRanges()
        {
            RangeClauses = new List<SummaryClauseRange<T>>();
        }

        public List<SummaryClauseRange<T>> RangeClauses { set; get; }
        public override bool HasCoverage => RangeClauses.Any();
        public bool Any() => HasCoverage;

        //public void Add(SummaryClauseRanges<T> newVal)
        //{
        //    Add(newVal.RangeClauses);
        //}

        //public void Add(IEnumerable<SummaryClauseRange<T>> newVals)
        //{
        //    foreach (var rg in newVals)
        //    {
        //        Add(rg);
        //    }
        //}

        public void Add(SummaryClauseRange<T> candidate)
        {
            if (Covers(candidate))
            {
                return;
            }

            AddRange(candidate);
        }

        public override bool Covers(ISummaryClause<T> candidate)
        {
            if (!HasCoverage)
            {
                return false;
            }

            if (candidate is SummaryClauseSingleValues<T> singleValues)
            {
                return singleValues.Values.All(sv => Covers(sv));
            }
            else if (candidate is SummaryClauseRanges<T> cRanges)
            {
                foreach (var rgCandidate in cRanges.RangeClauses)
                {
                    if (!RangeClauses.Any(rg => Covers(rgCandidate)))
                    {
                        return false;
                    }
                }
                return true;
            }
            else if (candidate is SummaryClauseRange<T> rangeCandidate)
            {
                return Covers(rangeCandidate);
            }
            return false;
        }

        private bool Covers(SummaryClauseRange<T> candidate)
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
            return RangeClauses.Any(rg => rg.Start.CompareTo(candidate) <= 0 && rg.End.CompareTo(candidate) >= 0);
        }

        public void Remove(List<SummaryClauseRange<T>> rangesToRemove)
        {
            foreach( var range in rangesToRemove)
            {
                Remove(range);
            }
        }

        public void Remove(SummaryClauseRange<T> rangeToRemove)
        {
            RangeClauses.Remove(rangeToRemove);
        }

        public override string ToString()
        {
            var result = string.Empty;
            foreach(var range in RangeClauses)
            {
                result = $"{result}Range={range.ToString()},";
            }
            if (result.Length > 0)
            {
                return result.Remove(result.Length - 1);
            }
            return string.Empty;
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
        //public void Add(Object o)
        //{
        //    if(o is SummaryClauseRanges<T> clauseRanges)
        //    {
        //        foreach(var range in clauseRanges.RangeClauses)
        //        {
        //            AddRange(range);
        //        }
        //    }
        //    else if (o is SummaryClauseRange<T> range)
        //    {
        //        AddRange(range);
        //    }
        //    else if (o is List<SummaryClauseRange<T>> list)
        //    {
        //        foreach (var rangeClause in list)
        //        {
        //            AddRange(rangeClause);
        //        }
        //    }
        //}

        private void AddRange(SummaryClauseRange<T> rangeClause)
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
