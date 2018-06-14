using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Rubberduck.Inspections.Concrete.UnreachableCaseInspection
{
    public enum VariableClauseTypes
    {
        Predicate,
        Value,
        Range,
        Is
    };

    public interface IExpressionFilter
    {
        void AddExpression(IRangeClauseExpression expression);
        void AddComparablePredicateFilter(string variable, string variableTypeName);
        bool HasFilters { get; }
        bool FiltersAllValues { get; }
    }

    public class ExpressionFilter<T> : IExpressionFilter where T : IComparable<T>
    {
        private struct PredicateValueExpression
        {
            private readonly int _hashCode;
            private readonly string _toString;

            public string LHS { private set; get; }
            public T RHS { private set; get; }
            public string OpSymbol { private set; get; }

            public PredicateValueExpression(string lhs, T rhs, string opSymbol)
            {
                LHS = lhs;
                RHS = rhs;
                OpSymbol = opSymbol;
                _toString = $"{LHS} {OpSymbol} {RHS}";
                _hashCode = _toString.GetHashCode();
            }

            public override string ToString() => _toString;
            public override int GetHashCode() => _hashCode;
            public override bool Equals(object obj)
            {
                if (!(obj is PredicateValueExpression expression))
                {
                    return false;
                }
                return _toString.Equals(expression.ToString());
            }
        }

        private readonly T _trueValue;
        private readonly T _falseValue;
        private readonly string _filterTypeName;
        private string _toString;

        public ExpressionFilter(StringToValueConversion<T> converter, string typeName)
        {
            Converter = converter;
            _filterTypeName = typeName;
            converter("True", typeName, out _trueValue);
            converter("False", typeName, out _falseValue);
        }

        private HashSet<IRangeClauseExpression> LikePredicates { get; } = new HashSet<IRangeClauseExpression>();

        private HashSet<PredicateValueExpression> ComparablePredicates { get; } = new HashSet<PredicateValueExpression>();

        private bool IsDirty { set; get; } = true;

        protected Dictionary<VariableClauseTypes, HashSet<string>> Variables { get; } = new Dictionary<VariableClauseTypes, HashSet<string>>()
        {
            [VariableClauseTypes.Is] = new HashSet<string>(),
            [VariableClauseTypes.Predicate] = new HashSet<string>(),
            [VariableClauseTypes.Range] = new HashSet<string>(),
            [VariableClauseTypes.Value] = new HashSet<string>(),
        };

        protected StringToValueConversion<T> Converter { set; get; } = null;

        protected HashSet<T> SingleValues { set; get; } = new HashSet<T>();

        protected HashSet<(T Start, T End)> Ranges { set; get; } = new HashSet<(T Start, T End)>();

        protected FilterLimits<T> Limits { get; } = new FilterLimits<T>();

        private Dictionary<string, IExpressionFilter> ComparablePredicateFilters { set; get; } = new Dictionary<string, IExpressionFilter>();

        private Dictionary<string, IExpressionFilter> ComparablePredicateFiltersInverse { set; get; } = new Dictionary<string, IExpressionFilter>();

        public void AddComparablePredicateFilter(string variable, string variableTypeName)
        {
            if (variable is null || variable.Length == 0 || variableTypeName is null || variableTypeName.Length == 0)
            {
                return;
            }

            if (!ComparablePredicateFilters.ContainsKey(variable))
            {
                ComparablePredicateFilters.Add(variable, ExpressionFilterFactory.Create(variableTypeName));
                ComparablePredicateFiltersInverse.Add(variable, ExpressionFilterFactory.Create(variableTypeName));
            }
        }

        public override bool Equals(object obj)
        {
            if (!(obj is ExpressionFilter<T> filter))
            {
                return false;
            }

            return Ranges.SetEquals(filter.Ranges)
                && SingleValues.SetEquals(filter.SingleValues)
                && ComparablePredicates.SetEquals(filter.ComparablePredicates)
                && LikePredicates.SetEquals(filter.LikePredicates)
                && this[VariableClauseTypes.Range].SetEquals(filter[VariableClauseTypes.Range])
                && this[VariableClauseTypes.Value].SetEquals(filter[VariableClauseTypes.Value])
                && this[VariableClauseTypes.Predicate].SetEquals(filter[VariableClauseTypes.Predicate])
                && this[VariableClauseTypes.Is].SetEquals(filter[VariableClauseTypes.Is])
                && Limits.Equals(filter.Limits);
        }

        public void SetExtents(T min, T max) => Limits.SetExtents(min, max);

        protected virtual bool TryGetMaximum(out T maximum) => Limits.TryGetMaximum(out maximum);

        protected virtual bool TryGetMinimum(out T minimum) => Limits.TryGetMinimum(out minimum);

        public void AddExpression(IRangeClauseExpression expression)
        {
            if (expression is null) { return; }

            try
            {
                switch (expression)
                {
                    case IsClauseExpression isClause:
                        expression.IsUnreachable =  !AddIsClause(isClause);
                        return;
                    case RangeOfValuesExpression rangeExpr:
                        expression.IsUnreachable = !AddRangeOfValuesExpression(rangeExpr);
                        return;
                    case ValueExpression valueExpr:
                        expression.IsUnreachable = !AddValueExpression(valueExpr);
                        return;
                    case UnaryExpression unaryExpr:
                        expression.IsUnreachable = !AddUnaryExpression(unaryExpr);
                        return;
                    case BinaryExpression binary:
                        expression.IsUnreachable = !AddBinaryExpression(binary);
                        return;
                }

            }
            catch (ArgumentException)
            {
                expression.IsMismatch = true;
            }
        }

        public virtual bool HasFilters => Ranges.Any()
                    || SingleValues.Any()
                    || Limits.Any()
                    || this[VariableClauseTypes.Value].Any()
                    || this[VariableClauseTypes.Range].Any()
                    || this[VariableClauseTypes.Is].Any()
                    || this[VariableClauseTypes.Predicate].Any()
                    || LikePredicates.Any()
                    || ComparablePredicates.Any();

        private bool AddLike(IRangeClauseExpression predicate)
        {
            if (FiltersTrueFalse) { return false; }

            return predicate.RHS.Equals("*") ? AddSingleValue(_trueValue) 
                : AddToContainer(LikePredicates, predicate);
        }

        protected bool AddComparablePredicate(string lhs, IRangeClauseExpression expression)
        {
            if (FiltersTrueFalse) { return false; }

            if (!Converter(expression.RHS, _filterTypeName, out T rhsVal))
            {
                throw new ArgumentOutOfRangeException($"Unable to convert {expression.RHS} to {typeof(T)}");
            }

            if (ComparablePredicateFilters.ContainsKey(lhs))
            {
                var positiveLogic = ComparablePredicateFilters[lhs];
                if (!positiveLogic.FiltersAllValues)
                {
                    IRangeClauseExpression predicateExpression = new IsClauseExpression(expression.RHSValue, expression.OpSymbol);
                    positiveLogic.AddExpression(predicateExpression);
                    if (positiveLogic.FiltersAllValues)
                    {
                        AddSingleValue(_trueValue);
                    }
                }

                var negativeLogic = ComparablePredicateFiltersInverse[lhs];
                if (!negativeLogic.FiltersAllValues)
                {
                    IRangeClauseExpression predicateExpressionInverse
                        = new IsClauseExpression(expression.RHSValue, RelationalInverse(expression.OpSymbol));
                    negativeLogic.AddExpression(predicateExpressionInverse);
                    if (negativeLogic.FiltersAllValues)
                    {
                        AddSingleValue(_falseValue);
                    }
                }
            }

            var predicate = new PredicateValueExpression(lhs, rhsVal, expression.OpSymbol);
            var matchingVariablesNames = ComparablePredicates.Where(pv => pv.LHS.CompareTo(predicate.LHS) == 0);

            if (!matchingVariablesNames.Any(cv => cv.Equals(predicate)))
            {
                AddToContainer(ComparablePredicates, predicate);
                return true;
            }
            return false;
        }

        protected bool AddSingleValue(T value) => AddToContainer(SingleValues, value);

        protected virtual bool AddValueRange((T Start, T End) range)
        {
            if (FiltersRange(range))
            {
                return false;
            }

            IsDirty = true;
            if (Limits.HasMinimum)
            {
                range = TrimStart(range, Limits.Minimum);
            }

            if (Limits.HasMaximum)
            {
                range = TrimEnd(range, Limits.Maximum);
            }

            if (!Ranges.Any())
            {
                Ranges.Add(range);
                return true;
            }
            else
            {
                var initial = Ranges.Count;
                RemoveRangesCoveredByRange(range);

                if (!TryMergeWithOverlappingRange(range))
                {
                    Ranges.Add(range);
                }
                return initial != Ranges.Count;
            }
        }

        private bool Covers((T Start, T End) existingRange, (T Start, T End) range)
            => existingRange.Start.CompareTo(range.Start) <= 0 && existingRange.End.CompareTo(range.End) >= 0;

        private bool Covers((T Start, T End) range, T value)
            => range.Start.CompareTo(value) <= 0 && range.End.CompareTo(value) >= 0;

        private (T Start, T End) TrimStart((T Start, T End) rangeToTrim, T value)
        {
            if (rangeToTrim.Start.CompareTo(value) < 0)
            {
                return (value, rangeToTrim.End);
            }
            return (rangeToTrim.Start, rangeToTrim.End);
        }

        private (T Start, T End) TrimEnd((T Start, T End) rangeToTrim, T value)
        {
            if (rangeToTrim.End.CompareTo(value) > 0)
            {
                return (rangeToTrim.Start, value);
            }
            return (rangeToTrim.Start, rangeToTrim.End);
        }

        private bool FiltersRange((T Start, T End) range)
        {
            return Limits.CoversRange(range)
                || Ranges.Any(rg => Covers(rg, range));
        }

        private bool FiltersValue(T value) =>
            SingleValues.Contains(value)
            || RangesCoversValue(value)
            || Limits.CoversValue(value);

        private bool RangesCoversValue(T value)
            => Ranges.Any(rg => rg.Start.CompareTo(value) <= 0 && rg.End.CompareTo(value) > 0);

        public virtual bool FiltersAllValues
        {
            get
            {
                if (Limits.HasMinAndMaxLimits)
                {
                    return Limits.Minimum.CompareTo(Limits.Maximum) > 0
                        || Ranges.Any(rg => Covers(rg, (Limits.Minimum, Limits.Maximum)))
                        || SingleValues.Any(sv => Limits.Minimum.CompareTo(Limits.Maximum) == 0 && sv.CompareTo(Limits.Minimum) == 0);
                }
                return false;
            }
        }

        protected bool FiltersTrueFalse => FiltersValue(_trueValue) && FiltersValue(_falseValue);

        private HashSet<string> this[VariableClauseTypes eType] => Variables[eType];

        private bool AddRangeOfValuesExpression(RangeOfValuesExpression rangeExpr)
        {
            if (rangeExpr.LHSValue.ParsesToConstantValue && rangeExpr.RHSValue.ParsesToConstantValue)
            {
                var (start, end) = ConvertRangeValues(rangeExpr.LHS, rangeExpr.RHS);

                //If an expression is X To Y where X > Y, then the Range Clause will never execute
                if (typeof(T) == typeof(bool) ? end.CompareTo(start) > 0 : start.CompareTo(end) > 0)
                {
                    rangeExpr.IsUnreachable = true;
                    return false;
                }
                return start.CompareTo(end) == 0 ?
                    AddSingleValue(start) : AddValueRange((start, end));
            }
            return AddToContainer(Variables[VariableClauseTypes.Range], rangeExpr.ToString());
        }

        private bool AddValueExpression(ValueExpression valueExpr)
        {
            if (valueExpr.LHSValue.ParsesToConstantValue)
            {
                if (Converter(valueExpr.LHS, _filterTypeName, out T result))
                {
                    return FiltersValue(result) ? false : AddSingleValue(result);
                }
                throw new ArgumentException();
            }
            return AddToContainer(Variables[VariableClauseTypes.Value], valueExpr.ToString());
        }

        private bool AddUnaryExpression(UnaryExpression unaryExpr)
        {
            if (FiltersTrueFalse) { return false; }

            if (unaryExpr.LHSValue.ParsesToConstantValue)
            {
                if (Converter(unaryExpr.LHS, _filterTypeName, out T result))
                {
                    return FiltersValue(result) ? false : AddSingleValue(result);
                }
                throw new ArgumentException();
            }
            return AddToContainer(Variables[VariableClauseTypes.Predicate], unaryExpr.ToString());
        }

        private bool AddBinaryExpression(BinaryExpression binary)
        {
            var opSymbol = binary.OpSymbol;
            if (FiltersTrueFalse && ParseTreeExpressionEvaluator.LogicOpsBinary.ContainsKey(binary.OpSymbol))
            {
                return false;
            }

            if (opSymbol.Equals(LogicSymbols.LIKE))
            {
                if (binary.RHS.Equals("*"))
                {
                    return AddToContainer(SingleValues, _trueValue);
                }
                return AddLike(binary);
            }

            if (!binary.LHSValue.ParsesToConstantValue && binary.RHSValue.ParsesToConstantValue)
            {
                if (!Converter(binary.RHS, _filterTypeName, out T value))
                {
                    throw new ArgumentException();
                }

                return AddComparablePredicate(binary.LHS, binary);
            }

            if (!binary.LHSValue.ParsesToConstantValue && !binary.RHSValue.ParsesToConstantValue)
            {
                return AddToContainer(Variables[VariableClauseTypes.Predicate], binary.ToString());
            }
            return false;
        }

        protected virtual bool AddIsClause(IsClauseExpression expression)
        {
            if (Converter(expression.LHS, _filterTypeName, out T value))
            {
                IsDirty = true;
                if (IsClauseAdders.ContainsKey(expression.OpSymbol))
                {
                    if (IsClauseAdders[expression.OpSymbol](this, value))
                    {
                        return true;
                    }
                }
                return false;
            }
            else
            {
                return AddToContainer(Variables[VariableClauseTypes.Is], expression.ToString());
            }
        }

        protected virtual bool AddMinimum(T value)
        {
            IsDirty = true;
            var result = Limits.SetMinimum(value);
            if(TryGetMinimum(out T min))
            {
                var newRanges = new HashSet<(T Start, T End)>();
                foreach ( var range in Ranges)
                {
                    newRanges.Add(TrimStart(range, min));
                }
                Ranges = newRanges;

                SingleValues.Where(sv => sv.CompareTo(min) < 0).ToList()
                    .ForEach(sv => SingleValues.Remove(sv));
            }
            return result;
        }

        protected virtual bool AddMaximum(T value)
        {
            IsDirty = true;
            var result =  Limits.SetMaximum(value);
            if (TryGetMaximum(out T max))
            {
                var newRanges = new HashSet<(T Start, T End)>();
                foreach (var range in Ranges)
                {
                    newRanges.Add(TrimEnd(range, max));
                }
                Ranges = newRanges;

                SingleValues.Where(sv => sv.CompareTo(max) > 0).ToList()
                    .ForEach(sv => SingleValues.Remove(sv));
            }
            return result;
        }

        protected void RemoveRangesCoveredByLimits()
        {
            var rangesToRemove = Ranges.Where(rg => Limits.CoversRange(rg));
            foreach(var range in rangesToRemove)
            {
                Ranges.Remove(range);
            }

        }

        protected void RemoveRangesCoveredByRange((T Start, T End) range)
        {
            var rangesToRemove = Ranges.Where(rg => Covers(range, rg)).ToList();
            rangesToRemove.ForEach(rtr => Ranges.Remove(rtr));
        }

        protected void RemoveSingleValuesCoveredByRanges()
            => SingleValues.Where(sv => Ranges.Any(rg => Covers(rg, sv)))
            .ToList().ForEach(tr => SingleValues.Remove(tr));

        protected bool TryMergeWithOverlappingRange((T Start, T End) range)
        {
            var endIsWithin = Ranges.Where(currentRange => Covers(currentRange, range.End));
            var startIsWithin = Ranges.Where(currentRange => Covers(currentRange, range.Start));

            var rangeIsAdded = false;
            if (endIsWithin.Any() || startIsWithin.Any())
            {
                rangeIsAdded = true;
                if (endIsWithin.Any())
                {
                    (T Start, T End) = endIsWithin.First();
                    Ranges.Remove(endIsWithin.First());
                    Ranges.Add((range.Start, End));
                }
                else
                {
                    (T Start, T End) = startIsWithin.First();
                    Ranges.Remove(startIsWithin.First());
                    Ranges.Add((Start, range.End));
                }
            }
            return rangeIsAdded;
        }

        private (T start, T end) ConvertRangeValues(string startVal, string endVal)
        {
            if (!Converter(startVal, _filterTypeName, out T start) || !Converter(endVal, _filterTypeName, out T end))
            {
                throw new ArgumentException();
            }
            return (start, end);
        }

        protected bool AddToContainer<K>(HashSet<K> container, K value)
        {
            if (container.Contains(value))
            {
                return false;
            }
            IsDirty = true;
            container.Add(value);
            return true;
        }

        private static Dictionary<string, Func<ExpressionFilter<T>, T, bool>> IsClauseAdders = new Dictionary<string, Func<ExpressionFilter<T>, T, bool>>()
        {
            [LogicSymbols.LT] = delegate (ExpressionFilter<T> rg, T value) { return rg.AddMinimum(value); },
            [LogicSymbols.LTE] = delegate (ExpressionFilter<T> rg, T value) { var min = rg.AddMinimum(value); var val = rg.AddSingleValue(value); return min || val; },
            [LogicSymbols.GT] = delegate (ExpressionFilter<T> rg, T value) { return rg.AddMaximum(value); },
            [LogicSymbols.GTE] = delegate (ExpressionFilter<T> rg, T value) { var max = rg.AddMaximum(value); var val = rg.AddSingleValue(value); return max || val; },
            [LogicSymbols.EQ] = delegate (ExpressionFilter<T> rg, T value) { return rg.AddSingleValue(value); },
            [LogicSymbols.NEQ] = delegate (ExpressionFilter<T> rg, T value) { var min = rg.AddMinimum(value); var max = rg.AddMaximum(value); return min || max; },
        };

        private string RelationalInverse(string opSymbol)
            => RelationalInverses.Keys.Contains(opSymbol) ? RelationalInverses[opSymbol] : opSymbol;

        private static Dictionary<string, string> RelationalInverses = new Dictionary<string, string>()
        {
            [LogicSymbols.LT] = LogicSymbols.GTE,
            [LogicSymbols.LTE] = LogicSymbols.GTE,
            [LogicSymbols.GT] = LogicSymbols.LTE,
            [LogicSymbols.GTE] = LogicSymbols.LT,
            [LogicSymbols.EQ] = LogicSymbols.NEQ,
            [LogicSymbols.NEQ] = LogicSymbols.EQ,
        };

        public override string ToString()
        {
            if (!IsDirty && _toString != null)
            {
                return _toString;
            }

            var descriptors = new HashSet<string>
            {
                Limits.ToString(),
                GetRangesDescriptor(),
                GetSinglesDescriptor(),
                BuildTypeDescriptor(Variables[VariableClauseTypes.Is].ToList(), "Is"),
                GetPredicatesDescriptor()
            };

            descriptors.Remove(string.Empty);

            var descriptor = new StringBuilder();
            for (var idx = 0; idx < descriptors.Count; idx++)
            {
                descriptor.Append(descriptors.ElementAt(idx));
            }

            _toString = descriptor.ToString();
            IsDirty = false;
            return _toString;
        }

        private string GetSinglesDescriptor()
        {
            var singles = SingleValues.Select(sv => sv.ToString()).ToList();
            singles.AddRange(this[VariableClauseTypes.Value]);
            return BuildTypeDescriptor(singles, "Values");
        }

        private string GetRangesDescriptor()
        {
            var values = Ranges.Select(rg => $"{rg.Start}:{rg.End}").ToList();
            values.AddRange(this[VariableClauseTypes.Range]);
            return BuildTypeDescriptor(values, "Ranges");
        }

        private string GetPredicatesDescriptor()
        {
            var result = new HashSet<string>();
            foreach (var val in ComparablePredicates)
            {
                result.Add(val.ToString());
            }

            foreach (var like in LikePredicates)
            {
                result.Add(like.ToString());
            }

            foreach (var predicate in Variables[VariableClauseTypes.Predicate])
            {
                result.Add(predicate.ToString());
            }
            return BuildTypeDescriptor(result.ToList(), "Predicates");
        }

        private string BuildTypeDescriptor<K>(List<K> values, string identifier)
        {
            if (!values.Any()) { return string.Empty; }

            StringBuilder series = new StringBuilder();
            values.ForEach(val => series.Append($"{val},"));
            return $"{identifier}({series.ToString().Substring(0, series.Length - 1)})";
        }
    }
}
