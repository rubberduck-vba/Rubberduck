using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Rubberduck.Inspections.Concrete
{
    public class SummaryClauseSingleValues<T> : SummaryClauseBase<T> where T : IComparable<T>
    {
        private HashSet<T> _values;
        private HashSet<bool> _booleanValues = new HashSet<bool>();

        public SummaryClauseSingleValues()
        {
            _values = new HashSet<T>();
            _booleanValues = new HashSet<bool>();
    }

        public HashSet<T> Values => _values;
        public HashSet<bool> ValuesBoolean => _booleanValues;
        public override bool HasCoverage => _values.Any() || _booleanValues.Any();

        public override string ToString()
        {
            var result = string.Empty;
            foreach (var val in Values)
            {
                result = $"{result}Single={val.ToString()},";
            }
            foreach (var val in _booleanValues)
            {
                result = $"{result}Bool={val.ToString()},";
            }
            return result.Length > 0 ? result.Remove(result.Length - 1) : string.Empty;
        }

        public void Clear()
        {
            Values.Clear();
        }

        public void Add(IEnumerable<T> singleValues)
        {
            foreach (var singleVal in singleValues)
            {
                Add(singleVal);
            }
        }

        public void Add(bool value)
        {
            _booleanValues.Add(value.ToString().Equals(bool.TrueString));
        }

        public void Add(T value)
        {
            if (ContainsBooleans)
            {
                _booleanValues.Add(value.ToString().Equals(bool.TrueString));
            }
            _values.Add(value);
        }

        public override bool Covers(ISummaryClause<T> candidate)
        {
            if (candidate is SummaryClauseSingleValues<T> singleValues)
            {
                if (HasCoverage)
                {
                    return singleValues.Values.All(sv => Covers(sv));
                }
            }
            return false;
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
                if (ContainsBooleans)
                {
                    return _booleanValues.Count;
                }
                return Values.Count;
            }
        }

        public bool Any()
        {
            return Values.Any() || _booleanValues.Any();
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

        public bool Covers(bool candidate)
        {
            if (HasCoverage)
            {
                return _booleanValues.Any(bv => bv == candidate);
            }
            return false;
        }

        public void Remove(List<bool> values)
        {
            foreach (var value in values)
            {
                Remove(value);
            }
        }

        public void Remove(List<T> values)
        {
            foreach (var value in values)
            {
                Remove(value);
            }
        }

        public void Remove(bool value)
        {
            //if (ContainsBooleans)
            //{
                _booleanValues.Remove(value);
            //}

            //Values.Remove(value);
        }

        public void Remove(T value)
        {
            //if (ContainsBooleans)
            //{
            //    _booleanValues.Remove(value.ToString().Equals(bool.TrueString));
            //}

            Values.Remove(value);
        }
    }
}
