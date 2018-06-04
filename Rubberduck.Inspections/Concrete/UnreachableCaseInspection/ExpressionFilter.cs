using Rubberduck.Parsing.Grammar;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

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
        IRangeClauseExpression AddExpression(IRangeClauseExpression expression);
        bool HasFilters { get; }
        bool FiltersAllValues { get; }
    }

    public class ExpressionFilter<T> : IExpressionFilter where T : IComparable<T>
    {
        private readonly T _trueValue;
        private readonly T _falseValue;

        public ExpressionFilter() { }

        public ExpressionFilter(StringToValueConversion<T> converter)
        {
            Converter = converter;
            converter("True", out _trueValue);
            converter("False", out _falseValue);
        }

        private HashSet<IRangeClauseExpression> LikePredicates { get; } = new HashSet<IRangeClauseExpression>();

        private HashSet<PredicateValueExpression<T>> ComparablePredicates { get; } = new HashSet<PredicateValueExpression<T>>();

        protected Dictionary<VariableClauseTypes, HashSet<string>> Variables { get; } = new Dictionary<VariableClauseTypes, HashSet<string>>()
        {
            [VariableClauseTypes.Is] = new HashSet<string>(),
            [VariableClauseTypes.Predicate] = new HashSet<string>(),
            [VariableClauseTypes.Range] = new HashSet<string>(),
            [VariableClauseTypes.Value] = new HashSet<string>(),
        };

        protected StringToValueConversion<T> Converter { set; get; } = null;

        protected HashSet<T> SingleValues { set; get; } = new HashSet<T>();

        protected HashSet<RangeValues<T>> Ranges { set; get; } = new HashSet<RangeValues<T>>();

        public override bool Equals(object obj)
        {
            if (!(obj is ExpressionFilter<T> filter))
            {
                return false;
            }

            return Ranges.SetEquals(filter.Ranges)
                && this[VariableClauseTypes.Range].SetEquals(filter[VariableClauseTypes.Range])
                && SingleValues.SetEquals(filter.SingleValues)
                && this[VariableClauseTypes.Value].SetEquals(filter[VariableClauseTypes.Value])
                && ComparablePredicates.SetEquals(filter.ComparablePredicates)
                && LikePredicates.SetEquals(filter.LikePredicates)
                && this[VariableClauseTypes.Predicate].SetEquals(filter[VariableClauseTypes.Predicate])
                && this[VariableClauseTypes.Is].SetEquals(filter[VariableClauseTypes.Is])
                && Limits.Equals(filter.Limits);
        }

        protected FilterLimits<T> Limits { get; } = new FilterLimits<T>();

        public void SetExtents(T min, T max) => Limits.SetExtents(min, max);

        protected virtual bool TryGetMaximum(out T maximum) => Limits.TryGetMaximum(out maximum);

        protected virtual bool TryGetMinimum(out T minimum) => Limits.TryGetMinimum(out minimum);

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

            DescriptorIsDirty = true;

            return predicate.RHS.Equals("*") ? AddSingleValue(_trueValue) 
                : AddToContainer(LikePredicates, predicate);
        }


        protected bool AddComparablePredicate(string lhs, string rhs, string opSymbol)
        {
            if (FiltersTrueFalse) { return false; }

            DescriptorIsDirty = true;

            var result = false;
            if (!Converter(rhs, out T rhsVal))
            {
                throw new ArgumentOutOfRangeException($"Unable to convert {rhs} to {typeof(T)}");
            }

            var predicate = new PredicateValueExpression<T>(lhs, rhsVal, opSymbol);

            var matchingVariables = ComparablePredicates
                .Where(pv => pv.LHS.CompareTo(predicate.LHS) == 0 && pv.OpSymbol.Equals(predicate.OpSymbol));
            if (matchingVariables.Any())
            {
                var current = matchingVariables.First();
                if (!current.Filters(predicate))
                {
                    ComparablePredicates.Remove(current);
                    ComparablePredicates.Add(predicate);
                    result = true;
                }
            }
            else
            {
                ComparablePredicates.Add(predicate);
                result = true;
            }
            return result;
        }

        protected bool AddSingleValue(T value)
        {
            return AddToContainer(SingleValues, value);
        }

        protected virtual bool AddValueRange(RangeValues<T> range)
        {
            DescriptorIsDirty = true;
            if (FiltersRange(range))
            {
                return false;
            }

            if (Limits.HasMinimum)
            {
                range = range.TrimStart(Limits.Minimum);
            }

            if (Limits.HasMaximum)
            {
                range = range.TrimEnd(Limits.Maximum);
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

        private bool FiltersRange(RangeValues<T> range)
        {
            return Limits.FiltersRange(range.Start, range.End)
                || Ranges.Any(rg => rg.Covers(range));
        }

        private bool FiltersValue(T value) =>
            SingleValues.Contains(value)
            || RangesFilterValue(value)
            || Limits.FiltersValue(value);

        private bool RangesFilterValue(T value)
            => Ranges.Any(rg => rg.Start.CompareTo(value) <= 0 && rg.End.CompareTo(value) > 0);

        public virtual bool FiltersAllValues
        {
            get
            {
                if (Limits.HasMinAndMaxLimits)
                {
                    return Limits.Minimum.CompareTo(Limits.Maximum) > 0
                        || Ranges.Any(rg => rg.Covers(new RangeValues<T>(Limits.Minimum, Limits.Maximum)))
                        || SingleValues.Any(sv => Limits.Minimum.CompareTo(Limits.Maximum) == 0 && sv.CompareTo(Limits.Minimum) == 0);
                }
                return false;
            }
        }

        protected bool FiltersTrueFalse => FiltersValue(_trueValue) && FiltersValue(_falseValue);

        private HashSet<string> this[VariableClauseTypes eType]
        {
            get => Variables.ContainsKey(eType) ? Variables[eType] : new HashSet<string>();
        }

        public IRangeClauseExpression AddExpression(IRangeClauseExpression expression)
        {
            var result = false;
            try
            {
                result = AddExpressionInternal(expression);
                expression.IsUnreachable = !result;
            }
            catch (ArgumentException)
            {
                expression.IsMismatch = true;
            }
            return expression;
        }

        private bool AddExpressionInternal(IRangeClauseExpression expression)
        {
            if (expression is IsClauseExpression isClause)
            {
                return AddIsClause(isClause);
            }
            if (expression is RangeValuesExpression)
            {
                if (expression.LHSValue.ParsesToConstantValue && expression.RHSValue.ParsesToConstantValue)
                {
                    var (start, end) = ConvertRangeValues(expression.LHS, expression.RHS);
                    var range = new RangeValues<T>(start, end);
                    if (range.IsUnreachable)
                    {
                        expression.IsUnreachable = true;
                        return false;
                    }
                    if (range.IsSingleValue)
                    {
                        return AddSingleValue(range.Start);
                    }
                    return AddValueRange(range);
                }
                else
                {
                    return AddToContainer(Variables[VariableClauseTypes.Range], expression.ToString());
                }
            }
            if (expression is ValueExpression)
            {
                if (expression.LHSValue.ParsesToConstantValue)
                {
                    if (Converter(expression.LHS, out T result))
                    {
                        if (FiltersValue(result))
                        {
                            return false;
                        }
                        return AddSingleValue(result);
                    }
                    else
                    {
                        throw new ArgumentException();
                    }
                }
                else
                {
                    return AddToContainer(Variables[VariableClauseTypes.Value], expression.ToString());
                }
            }
            if(expression is UnaryExpression)
            {
                if (FiltersTrueFalse) { return false; }

                if (expression.LHSValue.ParsesToConstantValue)
                {
                    if (Converter(expression.LHS, out T result))
                    {
                        if (FiltersValue(result))
                        {
                            return false;
                        }
                        return AddSingleValue(result);
                    }
                    else
                    {
                        throw new ArgumentException();
                    }
                }
                else
                {
                    return AddToContainer(Variables[VariableClauseTypes.Predicate], expression.ToString());
                }
            }
            if (expression is BinaryExpression binary)
            {
                if (FiltersTrueFalse) { return false; }

                if (binary.OpSymbol.Equals(LogicSymbols.LIKE))
                {
                    if (binary.RHS.Equals("*"))
                    {
                        return AddToContainer(SingleValues, _trueValue);
                    }
                    return AddLike(binary);
                }

                if (!expression.LHSValue.ParsesToConstantValue && expression.RHSValue.ParsesToConstantValue)
                {
                    if (!Converter(expression.RHS, out T value))
                    {
                        throw new ArgumentException();
                    }

                    DescriptorIsDirty = true;
                    var ptv = new PredicateValueExpression<T>(expression.LHS, value, expression.OpSymbol);
                    var currentVars = ComparablePredicates.Where(pv => pv.LHS.Equals(ptv.LHS) && pv.OpSymbol.Equals(ptv.OpSymbol));
                    var replace = false;
                    if (currentVars.Any())
                    {
                        //TODO: Need other Logic symbols here
                        var current = currentVars.First();
                        if (current.OpSymbol.Equals(LogicSymbols.LT))
                        {
                            replace = current.RHS.CompareTo(value) < 0;
                        }
                    }
                    else
                    {
                        ComparablePredicates.Add(ptv);
                        return true;
                    }
                    if (replace)
                    {
                        ComparablePredicates.Remove(currentVars.First());
                        ComparablePredicates.Add(ptv);
                    }
                    return replace;
                }
                else if (!expression.LHSValue.ParsesToConstantValue && !expression.RHSValue.ParsesToConstantValue)
                {
                    return AddToContainer(Variables[VariableClauseTypes.Predicate], binary.ToString());
                }
            }
            return false;
        }

        protected bool AddToContainer<K>(HashSet<K> container, K value)
        {
            if (container.Contains(value))
            {
                return false;
            }
            container.Add(value);
            return true;
        }

        private (T start, T end) ConvertRangeValues(string startVal, string endVal)
        {
            if (!Converter(startVal, out T start) || !Converter(endVal, out T end))
            {
                throw new ArgumentException();
            }
            return (start, end);
        }

        //private (bool hasValue, T value) TValue(IParseTreeValue ptValue)
        //{
        //    bool hasValue = Converter(ptValue.ToString(), out T value);
        //    return (hasValue, value);
        //}

        protected virtual bool AddIsClause(IsClauseExpression expression)
        {
            DescriptorIsDirty = true;
            if (Converter(expression.LHS, out T value))
            {
                if (IsClauseAdders.ContainsKey(expression.OpSymbol))
                {
                    if(IsClauseAdders[expression.OpSymbol](this, value))
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
            var result = Limits.SetMinimum(value);
            if(TryGetMinimum(out T min))
            {
                var newRanges = new HashSet<RangeValues<T>>();
                foreach( var range in Ranges)
                {
                    newRanges.Add(range.TrimStart(min));
                }
                Ranges = newRanges;

                SingleValues.Where(sv => sv.CompareTo(min) < 0).ToList()
                    .ForEach(sv => SingleValues.Remove(sv));
            }
            DescriptorIsDirty = true;
            return result;
        }

        protected virtual bool AddMaximum(T value)
        {
            var result =  Limits.SetMaximum(value);
            if (TryGetMaximum(out T max))
            {
                var newRanges = new HashSet<RangeValues<T>>();
                foreach (var range in Ranges)
                {
                    newRanges.Add(range.TrimEnd(max));
                }
                Ranges = newRanges;

                SingleValues.Where(sv => sv.CompareTo(max) > 0).ToList()
                    .ForEach(sv => SingleValues.Remove(sv));
            }
            DescriptorIsDirty = true;
            return result;
        }

        protected void RemoveRangesCoveredByLimits()
        {
            var rangesToRemove = Ranges.Where(rg => Limits.FiltersRange(rg.Start, rg.End));
            foreach(var range in rangesToRemove)
            {
                Ranges.Remove(range);
            }

        }

        protected void RemoveRangesCoveredByRange(RangeValues<T> range)
        {
            var rangesToRemove = Ranges.Where(rg => range.Covers(rg)).ToList();
            rangesToRemove.ForEach(rtr => Ranges.Remove(rtr));
        }

        protected void RemoveSingleValuesCoveredByRanges()
            => SingleValues.Where(sv => Ranges.Any(rg => rg.Covers(sv)))
            .ToList().ForEach(tr => SingleValues.Remove(tr));

        protected bool TryMergeWithOverlappingRange(RangeValues<T> range)
        {
            var endIsWithin = Ranges.Where(currentRange => currentRange.Covers(range.End));
            var startIsWithin = Ranges.Where(currentRange => currentRange.Covers(range.Start));

            var rangeIsAdded = false;
            if (endIsWithin.Any() || startIsWithin.Any())
            {
                rangeIsAdded = true;
                if (endIsWithin.Any())
                {
                    var original = endIsWithin.First();
                    Ranges.Remove(endIsWithin.First());
                    Ranges.Add(new RangeValues<T>(range.Start, original.End));
                }
                else
                {
                    var original = startIsWithin.First();
                    Ranges.Remove(startIsWithin.First());
                    Ranges.Add(new RangeValues<T>(original.Start, range.End));
                }
            }
            return rangeIsAdded;
        }

        private bool DescriptorIsDirty { set; get; }

        private string CachedToStringResult { set; get; }

        private bool StoreVariable(Dictionary<VariableClauseTypes, HashSet<string>> storage, VariableClauseTypes variableType, string value)
        {
            if (!storage.ContainsKey(variableType))
            {
                storage.Add(variableType, new HashSet<string>());
            }
            return AddToContainer(storage[variableType], value);
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

        public override string ToString()
        {
            if (!DescriptorIsDirty && CachedToStringResult != null)
            {
                return CachedToStringResult;
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

            CachedToStringResult = descriptor.ToString();
            DescriptorIsDirty = false;
            return CachedToStringResult;
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
