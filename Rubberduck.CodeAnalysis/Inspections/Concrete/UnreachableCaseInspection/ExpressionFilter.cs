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
        IRangeClauseExpression AddExpression(IRangeClauseExpression expression);
        bool HasFilters { get; }
        bool FiltersAllValues { get; }
    }

    public class ExpressionFilter<T> : IExpressionFilter where T : IComparable<T>
    {
        private readonly T _trueValue;
        private readonly T _falseValue;
        private readonly string _filterTypeName;

        public ExpressionFilter() { }

        public ExpressionFilter(StringToValueConversion<T> converter, string typeName)
        {
            Converter = converter;
            _filterTypeName = typeName;
            converter("True", _filterTypeName, out _trueValue);
            converter("False", _filterTypeName, out _falseValue);
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

        protected HashSet<(T Start, T End)> Ranges { set; get; } = new HashSet<(T Start, T End)>();

        protected FilterLimits<T> Limits { get; } = new FilterLimits<T>();

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

            if (!Converter(rhs, _filterTypeName, out T rhsVal))
            {
                throw new ArgumentOutOfRangeException($"Unable to convert {rhs} to {typeof(T)}");
            }

            var predicate = new PredicateValueExpression<T>(lhs, rhsVal, opSymbol);

            var matchingVariablesAndSymbol = ComparablePredicates
                .Where(pv => pv.LHS.CompareTo(predicate.LHS) == 0 && pv.OpSymbol.Equals(predicate.OpSymbol));

            if (matchingVariablesAndSymbol.Any(cv => cv.Equals(predicate)))
            {
                return false;
            }
            else
            {
                ComparablePredicates.Add(predicate);
                return true;
            }
        }

        protected bool AddSingleValue(T value)
        {
            return AddToContainer(SingleValues, value);
        }

        protected virtual bool AddValueRange((T Start, T End) range)
        {
            DescriptorIsDirty = true;
            if (FiltersRange(range))
            {
                return false;
            }

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

        protected bool FiltersTrueFalse => (FiltersValue(_trueValue) && FiltersValue(_falseValue))
            || ComparablePredicates.Where(cp => cp.OpSymbol.Equals(LogicSymbols.NEQ)).Count() > 1;

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
            switch (expression)
            {
                case IsClauseExpression isClause:
                    return AddIsClause(isClause);
                case RangeValuesExpression rangeExpr:
                    if (rangeExpr.LHSValue.ParsesToConstantValue && rangeExpr.RHSValue.ParsesToConstantValue)
                    {
                        var (start, end) = ConvertRangeValues(expression.LHS, expression.RHS);
                        if (typeof(T) == typeof(bool) ? end.CompareTo(start) > 0 : start.CompareTo(end) > 0)
                        {
                            rangeExpr.IsUnreachable = true;
                            return false;
                        }
                        return start.CompareTo(end) == 0 ?
                            AddSingleValue(start) : AddValueRange((start, end));
                    }
                    else
                    {
                        return AddToContainer(Variables[VariableClauseTypes.Range], expression.ToString());
                    }
                case ValueExpression valueExpr:
                    if (valueExpr.LHSValue.ParsesToConstantValue)
                    {
                        if (Converter(valueExpr.LHS, _filterTypeName, out T result))
                        {
                            return FiltersValue(result) ? false : AddSingleValue(result);
                        }
                        throw new ArgumentException();
                    }
                    else
                    {
                        return AddToContainer(Variables[VariableClauseTypes.Value], valueExpr.ToString());
                    }
                case UnaryExpression unaryExpr:
                    if (FiltersTrueFalse) { return false; }

                    if (unaryExpr.LHSValue.ParsesToConstantValue)
                    {
                        if (Converter(unaryExpr.LHS, _filterTypeName, out T result))
                        {
                            return FiltersValue(result) ? false : AddSingleValue(result);
                        }
                        throw new ArgumentException();
                    }
                    else
                    {
                        return AddToContainer(Variables[VariableClauseTypes.Predicate], expression.ToString());
                    }
                case BinaryExpression binary:
                    var opSymbol = binary.OpSymbol;
                    if (FiltersTrueFalse && ParseTreeExpressionEvaluator.LogicOpsBinary.ContainsKey(binary.OpSymbol)
                        && !(binary.OpSymbol.Equals(LogicSymbols.EQV) || binary.OpSymbol.Equals(LogicSymbols.IMP)))
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

                    if (!expression.LHSValue.ParsesToConstantValue && expression.RHSValue.ParsesToConstantValue)
                    {
                        if (!Converter(expression.RHS, _filterTypeName, out T value))
                        {
                            throw new ArgumentException();
                        }

                        DescriptorIsDirty = true;
                        var ptv = new PredicateValueExpression<T>(expression.LHS, value, expression.OpSymbol);
                        return AddComparablePredicate(expression.LHS, expression.RHS, expression.OpSymbol);
                        //var currentVars = ComparablePredicates.Where(pv => pv.LHS.Equals(expression.LHS) && pv.OpSymbol.Equals(opSymbol));
                        //var replaceCurrent = false;
                        //var add = false;

                        //if (currentVars.Any(cv => cv.Equals(ptv)))
                        //{
                        //    return false;
                        //}
                        ////TODO: Need to handle all the commented (below) logic types
                        ////TODO: Put the comparisons here in a loop like the comparison for Copy/Paste above

                        //if (!currentVars.Any())
                        //{
                        //    ComparablePredicates.Add(ptv);
                        //    return true;
                        //}
                        //else
                        //{
                        //    foreach (var cv in currentVars)
                        //    {
                        //        //DONEpublic static string EQ => _equalTo ?? LoadSymbols(VBAParser.EQ);
                        //        //DONEpublic static string NEQ => "<>";
                        //        //DONEpublic static string LT => _lessThan ?? LoadSymbols(VBAParser.LT);
                        //        //DONEpublic static string LTE => "<=";
                        //        //DONEpublic static string GT => _greaterThan ?? LoadSymbols(VBAParser.GT);
                        //        //DONEpublic static string GTE => ">=";
                        //        //public static string AND => Tokens.And;
                        //        //public static string OR => Tokens.Or;
                        //        //public static string XOR => Tokens.XOr;
                        //        //public static string NOT => Tokens.Not;
                        //        //public static string EQV => Tokens.Eqv;
                        //        //public static string IMP => Tokens.Imp;
                        //        //public static string LIKE => Tokens.Like;

                        //        var current = currentVars.First();
                        //        if (current.OpSymbol.Equals(LogicSymbols.LT))
                        //        {
                        //            replaceCurrent = current.RHS.CompareTo(value) < 0;
                        //        }
                        //        else if (current.OpSymbol.Equals(LogicSymbols.LTE))
                        //        {
                        //            replaceCurrent = current.RHS.CompareTo(value) <= 0;
                        //        }
                        //        else if (current.OpSymbol.Equals(LogicSymbols.GT))
                        //        {
                        //            replaceCurrent = current.RHS.CompareTo(value) > 0;
                        //        }
                        //        else if (current.OpSymbol.Equals(LogicSymbols.GTE))
                        //        {
                        //            replaceCurrent = current.RHS.CompareTo(value) >= 0;
                        //        }
                        //        else if (current.OpSymbol.Equals(LogicSymbols.NEQ))
                        //        {
                        //            add = true;
                        //        }
                        //        else if (current.OpSymbol.Equals(LogicSymbols.EQ))
                        //        {
                        //            add = true;
                        //        }
                        //    }
                        //}
                        ////else
                        ////{
                        ////    ComparablePredicates.Add(ptv);
                        ////    return true;
                        ////}

                        //if (replaceCurrent)
                        //{
                        //    ComparablePredicates.Remove(currentVars.First());
                        //    ComparablePredicates.Add(ptv);
                        //}
                        //if (add)
                        //{
                        //    ComparablePredicates.Add(ptv);
                        //}
                        //return replaceCurrent;
                    }
                    else if (!expression.LHSValue.ParsesToConstantValue && !expression.RHSValue.ParsesToConstantValue)
                    {
                        return AddToContainer(Variables[VariableClauseTypes.Predicate], binary.ToString());
                    }
                    return false;
                default:
                    return false;
            }
        }
        private bool AddVariablePredicate(IRangeClauseExpression expression)
        {
            //x < 65, if x > <val>, where <val> < 65 exists or x = <val> and x > <val> where <val> = 65 exists,
            //then T/F is covered
            //LT,LTE
            //if (LT)RHS < MinLimit - unreachable 
            //if (LTE)RHS is in SingleValues, and RHS < MinLimit - unreachable
            return false;
        }
        //private bool AddVariableRange(IRangeClauseExpression expression)
        //{
        //    if (Variables[VariableClauseTypes.Range].Contains(expression.ToString()))
        //    {
        //        return false;
        //    }

        //    if (expression.LHSValue.ParsesToConstantValue)
        //    {
        //        var 
        //    }
        //    return AddToContainer(Variables[VariableClauseTypes.Range], expression.ToString());
        //}

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
            if (!Converter(startVal, _filterTypeName, out T start) || !Converter(endVal, _filterTypeName, out T end))
            {
                throw new ArgumentException();
            }
            return (start, end);
        }

        protected virtual bool AddIsClause(IsClauseExpression expression)
        {
            DescriptorIsDirty = true;
            if (Converter(expression.LHS, _filterTypeName, out T value))
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
                var newRanges = new HashSet<(T Start, T End)>();
                foreach ( var range in Ranges)
                {
                    newRanges.Add(TrimStart(range, min));
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
                var newRanges = new HashSet<(T Start, T End)>();
                foreach (var range in Ranges)
                {
                    newRanges.Add(TrimEnd(range, max));
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
            var rangesToRemove = Ranges.Where(rg => Limits.CoversRange(rg)); //.Start, rg.End));
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
