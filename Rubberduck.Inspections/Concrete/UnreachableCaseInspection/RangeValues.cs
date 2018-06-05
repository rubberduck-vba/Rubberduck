using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Rubberduck.Inspections.Concrete.UnreachableCaseInspection
{
    public struct RangeValues<T> where T : IComparable<T>
    {
        public RangeValues(T start, T end)
        {
            Start = start; // start.CompareTo(end) > 0 ? end : start;
            End = end; // start.CompareTo(end) > 0 ? start : end;
        }

        public T Start { private set; get; }

        public T End { private set; get; }

        //In VBA-land True is less than False, not so in C#
        public bool IsUnreachable => 
            typeof(T) == typeof(bool) ? End.CompareTo(Start) > 0 : Start.CompareTo(End) > 0;

        public bool IsSingleValue => Start.CompareTo(End) == 0;

        public bool Covers(T value) 
            => Start.CompareTo(value) <= 0 && End.CompareTo(value) >= 0;

        public bool Covers(RangeValues<T> range)
            => Start.CompareTo(range.Start) <= 0 && End.CompareTo(range.End) >= 0;

        public RangeValues<T> TrimStart(T value)
        {
            if (Start.CompareTo(value) < 0)
            {
                return new RangeValues<T>(value, End);
            }
            return new RangeValues<T>(Start, End);
        }

        public RangeValues<T> TrimEnd(T value)
        {
            if (End.CompareTo(value) > 0)
            {
                return new RangeValues<T>(Start, value);
            }
            return new RangeValues<T>(Start, End);
        }

        //public override bool Equals(object obj)
        //{
        //    if (!(obj is RangeValues<T> rangeObj))
        //    {
        //        return false;
        //    }
        //    return rangeObj.Start.CompareTo(Start) == 0 && rangeObj.End.CompareTo(End) == 0;
        //}

        //public override int GetHashCode()
        //{
        //    return _hashCode;
        //}

        public override string ToString()
        {
            return $"{Start}:{End}";
        }
    }
}
