using System;

namespace Rubberduck.CodeAnalysis.Inspections.Concrete.UnreachableCaseEvaluation
{
    internal struct Limit<T> where T : IComparable<T>
    {
        public bool HasValue;
        public T Value;

        public static bool operator >(Limit<T> LHS, Limit<T> RHS)
        {
            if (!LHS.HasValue || !RHS.HasValue)
            {
                return false;
            }
            return LHS.Value.CompareTo(RHS.Value) > 0;
        }

        public static bool operator <(Limit<T> LHS, Limit<T> RHS)
        {
            if (!LHS.HasValue || !RHS.HasValue)
            {
                return false;
            }
            return LHS.Value.CompareTo(RHS.Value) < 0;
        }

        public static bool operator ==(Limit<T> LHS, Limit<T> RHS)
        {
            if (LHS.HasValue && RHS.HasValue)
            {
                return LHS.Value.CompareTo(RHS.Value) == 0;
            }
            return !LHS.HasValue && !RHS.HasValue;
        }

        public static bool operator !=(Limit<T> LHS, Limit<T> RHS) => !(LHS == RHS);

        public static bool operator >(Limit<T> LHS, T RHS) => LHS.HasValue && LHS.Value.CompareTo(RHS) > 0;

        public static bool operator <(Limit<T> LHS, T RHS) => LHS.HasValue && LHS.Value.CompareTo(RHS) < 0;

        public static bool operator >=(Limit<T> LHS, Limit<T> RHS) => LHS == RHS || LHS > RHS;

        public static bool operator <=(Limit<T> LHS, Limit<T> RHS) => LHS == RHS || LHS < RHS;

        public override int GetHashCode() => HasValue ? Value.GetHashCode() : base.GetHashCode();

        public override bool Equals(object obj)
        {
            if (!(obj is Limit<T> limit))
            {
                return false;
            }
            return this == limit;
        }

        public override string ToString()
        {
            return HasValue ? Value.ToString() : "NaN";
        }
    }

    internal class FilterLimits<T> where T : IComparable<T>
    {
        private Limit<T> _min;
        private Limit<T> _max;

        public FilterLimits()
        {
            _min = new Limit<T>()
            {
                Value = default,
                HasValue = false
            };

            _max = new Limit<T>()
            {
                Value = default,
                HasValue = false
            };
        }

        public T Maximum => _max.Value;

        public T Minimum => _min.Value;

        public Limit<T> MinimumExtent { set; get; } = default;

        public Limit<T> MaximumExtent { set; get; } = default;

        public void SetExtents(T min, T max)
        {
            MinimumExtent = new Limit<T>() { Value = min, HasValue = true };
            MaximumExtent = new Limit<T>() { Value = max, HasValue = true };
            if (!HasMinimum)
            {
                SetMinimum(min);
            }

            if (!HasMaximum)
            {
                SetMaximum(max);
            }
        }

        public bool SetMinimum(T min)
        {
            bool setNewValue;
            if (_min.HasValue)
            {
                setNewValue = min.CompareTo(_min.Value) > 0;
                _min.Value = setNewValue ? min : _min.Value;
            }
            else
            {
                setNewValue = true;
                _min.Value = min;
                _min.HasValue = true;
            }
            return setNewValue;
        }

        public bool SetMaximum(T max)
        {
            bool setNewValue;
            if (_max.HasValue)
            {
                setNewValue = max.CompareTo(_max.Value) < 0;
                _max.Value = setNewValue ? max : _max.Value;
            }
            else
            {
                setNewValue = true;
                _max.Value = max;
                _max.HasValue = true;
            }
            return setNewValue;
        }

        public bool Any() => _min.HasValue || _max.HasValue;

        public bool HasMinAndMaxLimits => _min.HasValue && _max.HasValue;

        public bool HasMinimum => _min.HasValue;

        public bool HasMaximum => _max.HasValue;

        public bool TryGetMaximum(out T maximum)
        {
            maximum = HasMaximum ? Maximum : default;
            return HasMaximum;
        }

        public bool TryGetMinimum(out T minimum)
        {
            minimum = HasMinimum ? Minimum : default;
            return HasMinimum;
        }

        public bool FiltersValue(T value) => _min > value || _max < value;

        public bool FiltersRange(T Start, T End) => _min > End || _max < Start;

        public override bool Equals(object obj)
        {
            if (!(obj is FilterLimits<T> filter))
            {
                return false;
            }

            if (HasMinimum && filter.HasMinimum)
            {
                if (Minimum.CompareTo(filter.Minimum) != 0)
                {
                    return false;
                }
            }
            else if (HasMinimum || filter.HasMinimum)
            {
                return false;
            }

            if (HasMaximum && filter.HasMaximum)
            {
                if (Maximum.CompareTo(filter.Maximum) != 0)
                {
                    return false;
                }
            }
            else if (HasMaximum || filter.HasMaximum)
            {
                return false;
            }
            return true;
        }

        public override int GetHashCode() => VBEditor.HashCode.Compute(Minimum, Maximum);

        public override string ToString()
        {
            var minString = string.Empty;
            var maxString = string.Empty;
            if (_min.HasValue)
            {
                minString = MinimumExtent.HasValue && _min == MinimumExtent
                    ? $"Min(typeMin)" : $"Min({_min})";
            }

            if (_max.HasValue)
            {
                maxString = MaximumExtent.HasValue && _max == MaximumExtent
                    ? $"Max(typeMax)" : $"Max({_max})";
            }
            return $"{minString}{maxString}";
        }
    }
}
