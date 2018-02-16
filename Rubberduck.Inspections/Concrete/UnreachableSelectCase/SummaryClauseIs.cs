using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Rubberduck.Inspections.Concrete
{
    public class SummaryClauseIsLT<T> : SummaryClauseIsBase<T> where T : IComparable<T>
    {
        public SummaryClauseIsLT() : base(true) { }
        public SummaryClauseIsLT(T value) : base(value, true) { }

        public override string ToString()
        {
            return HasCoverage ? $"IsLT={Value}" : string.Empty;
        }

        public void Add(SummaryClauseIsLT<T> candidate)
        {
            if (candidate.HasCoverage)
            {
                Value = candidate.Value;
            }
        }

        public override bool Covers(T candidate)
        {
            if (ContainsBooleans)
            {
                return false;
            }
            if (HasCoverage)
            {
                return _value.CompareTo(candidate) > 0;
            }
            if (HasExtents)
            {
                return _extentMin.CompareTo(candidate) > 0;
            }
            return false;
        }
    }

    public class SummaryClauseIsGT<T> : SummaryClauseIsBase<T> where T : IComparable<T>
    {
        public SummaryClauseIsGT() : base(false) { }
        public SummaryClauseIsGT(T value) : base(value, false) { }

        public override string ToString()
        {
            return HasCoverage ? $"IsGT={Value}" : string.Empty;
        }

        public void Add(SummaryClauseIsGT<T> candidate)
        {
            if (candidate.HasCoverage)
            {
                Value = candidate.Value;
            }
        }

        public override bool Covers(T candidate)
        {
            if (HasCoverage)
            {
                return _value.CompareTo(candidate) < 0;
            }
            if (HasExtents)
            {
                return _extentMax.CompareTo(candidate) < 0;
            }
            return false;
        }
    }

    public class SummaryClauseIsBase<T> : SummaryClauseSingleValueBase<T> where T : IComparable<T>
    {
        protected T _value;
        protected bool _hasValue;
        protected T _extentMin;
        protected T _extentMax;
        protected bool _hasExtents;

        public SummaryClauseIsBase(T value, bool isLT)
        {
            IsLTClause = isLT;
            Value = value;
        }

        public SummaryClauseIsBase(bool isLT)
        {
            IsLTClause = isLT;
            _hasValue = false;
            _value = default;
        }

        public bool HasValue => _hasValue;
        public bool HasExtents => _hasExtents;

        public void ApplyExtents(T min, T max)
        {
            _hasExtents = true;
            _extentMin = min;
            _extentMax = max;
            if (_hasValue && IsLTClause)
            {
                if (_value.CompareTo(min) < 0)
                {
                    _value = min;
                    _hasValue = false;
                }
            }
            if (_hasValue && !IsLTClause)
            {
                if (_value.CompareTo(max) > 0)
                {
                    _value = max;
                    _hasValue = false;
                }
            }
        }

        public override bool Covers(T candidate) { return false; }

        public long? AsIntegerNumber
        {
            get
            {
                long? result = null;
                if (ContainsIntegerNumbers )
                {
                    if (HasValue)
                    {
                        result = long.Parse(Value.ToString());
                    }
                    else if (HasExtents)
                    {
                        result = IsLTClause ? long.Parse(_extentMin.ToString()) : long.Parse(_extentMax.ToString());
                    }
                }
                return result;
            }
        }

        public override void Add(T value)
        {
            Value = value;
        }

        public override bool HasCoverage
        {
            get
            {
                if (HasValue)
                {
                    if (IsLTClause)
                    {
                        return HasExtents ? _value.CompareTo(_extentMin) != 0 : true;
                    }
                    else
                    {
                        return HasExtents ? _value.CompareTo(_extentMax) != 0 : true;
                    }
                }
                return false;
            }
        }

        public bool IsLTClause { set; get; }

        public void ClearIfCoveredBy(SummaryClauseIsBase<T> isClause)
        {
            if (isClause.Covers(Value))
            {
                Clear();
            }
        }

        public void Clear()
        {
            _value = default;
            _hasValue = false;
        }

        public T Value
        {
            set
            {
                if (ContainsBooleans)
                {
                    //TODO: introduce Truth table of observed behavior and write to SingleValues
                    return;
                }

                if (IsLTClause)
                {
                    if (HasValue)
                    {
                        _value = value.CompareTo(_value) > 0 ? value : _value;
                    }
                    else if (HasExtents)
                    {
                        _value = value.CompareTo(_extentMin) > 0 ? value : _extentMin;
                    }
                    else
                    {
                        _value = value;
                    }
                    _hasValue = true;
                }
                else
                {
                    if (HasValue)
                    {
                        _value = value.CompareTo(_value) < 0 ? value : _value;
                    }
                    else if (HasExtents)
                    {
                        _value = value.CompareTo(_extentMax) < 0 ? value : _extentMax;
                    }
                    else
                    {
                        _value = value;
                    }
                    _hasValue = true;
                }
            }

            get => _value;
        }
    }
}
