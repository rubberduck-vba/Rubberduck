using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Rubberduck.Inspections.Concrete
{
    public interface ISummaryClause<T> where T : System.IComparable<T>
    {
        bool Covers(ISummaryClause<T> candidate);
        bool Covers(T candidate);
        //void Add(ISummaryClause<T> candidate);
        bool HasCoverage { get; }
    }

    public abstract class SummaryClauseBase<T> : ISummaryClause<T> where T : System.IComparable<T>
    {
        public abstract bool HasCoverage { get; }
        public bool IsEmpty => !HasCoverage;
        //public abstract void Add(ISummaryClause<T> candidate);
        public abstract bool Covers(ISummaryClause<T> candidate);
        public abstract bool Covers(T candidate);
        public bool ContainsBooleans => typeof(T) == typeof(bool);
        public bool ContainsIntegerNumbers => typeof(T) == typeof(long) || typeof(T) == typeof(Int32) || typeof(T) == typeof(byte);
    }
}
