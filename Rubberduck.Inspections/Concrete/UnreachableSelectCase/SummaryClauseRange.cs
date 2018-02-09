using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Rubberduck.Inspections.Concrete
{
    public class SummaryClauseRange<T> : SummaryClauseBase<T> where T : System.IComparable<T>
    {
        private bool _hasStart;
        private bool _hasEnd;
        T _start;
        T _end;

        public SummaryClauseRange()
        {
            _hasStart = false;
            _hasEnd = false;
        }

        public SummaryClauseRange(T start, T end)
        {
            _hasStart = false;
            _hasEnd = false;
            if (start.CompareTo(end) <= 0)
            {
                Start = start;
                End = end;
            }
            else
            {
                Start = end;
                End = start;
            }
        }

        public override bool HasCoverage => _hasStart && _hasEnd;

        public T Start
        {
            get { return _start; }
            set
            {
                _start = value;
                _hasStart = true;
            }
        }

        public T End
        {
            get { return _end; }
            set
            {
                _end = value;
                _hasEnd = true;
            }
        }

        public override bool Covers(ISummaryClause<T> candidate)
        {
            if (!(candidate.HasCoverage || HasCoverage))
            {
                return true;
            }

            else if (candidate is SummaryClauseRanges<T> cRanges)
            {
                return Covers(cRanges);
            }
            else if (candidate is SummaryClauseRange<T> cRange)
            {
                return Covers(cRange.Start) && Covers(cRange.End);
            }
            return false;
        }

        public List<long> AsIntegerNumbers
        {
            get
            {
                var results = new List<long>();
                if (ContainsIntegerNumbers && HasCoverage)
                {
                    long startVal = long.Parse(Start.ToString());
                    long endVal = long.Parse(End.ToString());
                    for (var val = startVal; val <= endVal; val++)
                    {
                        results.Add(val);
                    }
                }
                return results;
            }
        }

        public override bool Covers(T candidate)
        {
            return Start.CompareTo(candidate) <= 0 && End.CompareTo(candidate) >= 0;
        }
        public override string ToString()
        {
            return $"{Start}:{End}";
        }

        public bool IsAdjacent(SummaryClauseRange<T> range)
        {
            if (!ContainsIntegerNumbers)
            {
                return false;
            }
            long thisStart = long.Parse(Start.ToString());
            long thisEnd = long.Parse(End.ToString());
            long testStart = long.Parse(range.Start.ToString());
            long testEnd = long.Parse(range.End.ToString());
            return testEnd == thisStart - 1 || testStart == thisEnd + 1;
        }

        public bool Overlaps(SummaryClauseRange<T> range)
        {
            if (Covers(range))
            {
                return true;
            }

            return End.CompareTo(range.End) > 0 && Start.CompareTo(range.Start) >= 0 && Start.CompareTo(range.End) <= 0
                || Start.CompareTo(range.Start) < 0 && End.CompareTo(range.End) <= 0 && End.CompareTo(range.Start) >= 0;
        }

        public void AppendRange(SummaryClauseRange<T> rangeToAppend)
        {
            long thisStart = long.Parse(Start.ToString());
            long thisEnd = long.Parse(End.ToString());
            long testStart = long.Parse(rangeToAppend.Start.ToString());
            long testEnd = long.Parse(rangeToAppend.End.ToString());
            if(testEnd == thisStart - 1)
            {
                Start = rangeToAppend.Start;
            }
            else if(testStart == thisEnd + 1)
            {
                End = rangeToAppend.End;
            }
        }

        public void RemoveOverlap(SummaryClauseRange<T> overlapRange)
        {
            if(End.CompareTo(overlapRange.End) > 0 && Start.CompareTo(overlapRange.Start) >= 0 && Start.CompareTo(overlapRange.End) <= 0)
            {
                Start = overlapRange.Start;
            }
            else if(Start.CompareTo(overlapRange.Start) < 0 && End.CompareTo(overlapRange.End) <= 0 && End.CompareTo(overlapRange.Start) >= 0)
            {
                End = overlapRange.End;
            }
        }
    }
}
