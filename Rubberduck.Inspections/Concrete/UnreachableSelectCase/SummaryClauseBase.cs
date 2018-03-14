using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Rubberduck.Inspections.Concrete.UnreachableSelectCase
{
    public interface ISummaryClause<T> where T : System.IComparable<T>
    {
        bool Covers(T candidate);
        bool HasCoverage { get; }
        T TrueValue { set; get; }
        T FalseValue { set; get; }
    }

    public interface ISummaryClauseSingleValues<T> : ISummaryClause<T> where T : System.IComparable<T>
    {
        void Add(T value);
    }

    public abstract class SummaryClauseBase<T> : ISummaryClause<T> where T : System.IComparable<T>
    {
        public abstract bool Covers(T candidate);
        public abstract bool HasCoverage { get; }
        public T TrueValue { set; get; }
        public T FalseValue { set; get; }

        public bool IsEmpty => !HasCoverage;
        public bool ContainsBooleans => typeof(T) == typeof(bool);
        public bool ContainsIntegerNumbers => typeof(T) == typeof(long) || typeof(T) == typeof(Int32) || typeof(T) == typeof(byte);
    }

    public abstract class SummaryClauseSingleValueBase<T> : ISummaryClauseSingleValues<T> where T : System.IComparable<T>
    {
        public abstract bool Covers(T candidate);
        public abstract bool HasCoverage { get; }
        public T TrueValue { set; get; }
        public T FalseValue { set; get; }

        public abstract void Add(T value);

        public bool IsEmpty => !HasCoverage;
        public bool ContainsBooleans => typeof(T) == typeof(bool);
        public bool ContainsIntegerNumbers => typeof(T) == typeof(long) || typeof(T) == typeof(Int32) || typeof(T) == typeof(byte);
    }
}
