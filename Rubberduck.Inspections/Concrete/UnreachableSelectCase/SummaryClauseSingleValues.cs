using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Rubberduck.Inspections.Concrete.UnreachableSelectCase
{
    public class SummaryClauseSingleValues<T> : SummaryClauseSingleValueBase<T> where T : IComparable<T>
    {
        private HashSet<T> _values;

        public SummaryClauseSingleValues(Func<IUnreachableCaseInspectionValue,T> tConverter) :base(tConverter)
        {
            _values = new HashSet<T>();
        }

        public HashSet<T> Values => _values;
        public override bool HasCoverage => _values.Any();

        public override string ToString()
        {
            const string prefix = "Single=";
            var result = prefix;
            foreach (var val in Values)
            {
                result = $"{result}{val.ToString()},";
            }
            if (result.Equals(prefix))
            {
                return string.Empty;
            }
            return  result.Length > 0 ? result.Remove(result.Length - 1) : string.Empty;
        }

        public override void Add(T value)
        {
            _values.Add(value);
        }

        public void Add(SummaryClauseSingleValues<T> singleVals)
        {
            foreach (var singleVal in singleVals.Values)
            {
                Add(singleVal);
            }
        }

        public void Add(IEnumerable<T> singleValues)
        {
            foreach (var singleVal in singleValues)
            {
                Add(singleVal);
            }
        }

        public bool Covers(List<long> integerValues)
        {
            if (ContainsIntegerNumbers && HasCoverage)
            {
                return integerValues.All(val => AsIntegerValues.Contains(val));
            }
            return false;
        }

        public int Count
        {
            get
            {
                return Values.Count;
            }
        }

        public List<long> AsIntegerValues
        {
            get
            {
                var results = new List<long>();
                if(ContainsIntegerNumbers)
                {
                    foreach( var val in Values)
                    {
                        results.Add(long.Parse(val.ToString()));
                    }
                }
                return results;
            }
        }

        public override bool Covers(T candidate)
        {
            if (HasCoverage)
            {
                return Values.Any(rg => rg.CompareTo(candidate) == 0);
            }
            return false;
        }

        public void RemoveIfCoveredBy(SummaryClauseSingleValues<T> singleValues)
        {
            List<T> toRemove = new List<T>();
            toRemove = Values.Where(sv => singleValues.Covers(sv)).ToList();

            Remove(toRemove);
        }

        public void Remove(List<T> values)
        {
            foreach (var value in values)
            {
                Remove(value);
            }
        }

        public void Remove(T value)
        {
            Values.Remove(value);
        }
    }
}
