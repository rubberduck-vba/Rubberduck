using Rubberduck.Parsing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Rubberduck.Inspections.Concrete.UnreachableCaseInspection
{
    public interface IRangeClauseFilter
    {
        bool ContainsFilters { get; }
        bool FiltersAllValues { get; }
        string TypeName { set; get; }
        IRangeClauseFilter FilterUnreachableClauses(IRangeClauseFilter filter);
        void Add(IRangeClauseFilter filter);
        void AddValueRange(IParseTreeValue inputStartValue, IParseTreeValue inputEndValue);
        void AddIsClause(IParseTreeValue value, string operatorSymbol);
        void AddSingleValue(IParseTreeValue value);
        void AddRelationalOperator(IParseTreeValue value);
    }

    public interface IRangeClauseFilterTestSupport<T>
    {
        bool TryGetIsLessThanValue(out T isLessThanValue);
        bool TryGetIsGreaterThanValue(out T isGreaterThanValue);
        HashSet<T> SingleValues { get; }
    }

    public class RangeClauseFilter<T> : IRangeClauseFilter, IRangeClauseFilterTestSupport<T> where T : IComparable<T>
    {
        private readonly IParseTreeValueFactory _valueFactory;
        private readonly IRangeClauseFilterFactory _filterFactory;
        private readonly TryConvertParseTreeValue<T> _valueConverter;
        private readonly T _trueValue;
        private readonly T _falseValue;

        private readonly Dictionary<string, List<T>> _isClause;

        private bool _hasExtents;
        private T _minExtent;
        private T _maxExtent;
        private string _cachedDescriptor;
        private bool _descriptorIsDirty;

        public RangeClauseFilter(string typeName, IParseTreeValueFactory valueFactory, IRangeClauseFilterFactory filterFactory, TryConvertParseTreeValue<T> tConverter)
        {
            _valueFactory = valueFactory;
            _filterFactory = filterFactory;
            _valueConverter = tConverter;

            _isClause = new Dictionary<string, List<T>>();
            _hasExtents = false;
            _falseValue = ConvertToContainedGeneric(false);
            _trueValue = ConvertToContainedGeneric(true);
            TypeName = typeName;
            _cachedDescriptor = string.Empty;
            _descriptorIsDirty = true;
        }

        public HashSet<T> SingleValues { get; } = new HashSet<T>();

        private List<(T Start, T End)> RangeValues { get; } = new List<(T Start, T End)>();
        private HashSet<string> VariableRanges { get; } = new HashSet<string>();
        private HashSet<string> RelationalOperators { get; } = new HashSet<string>();
        private HashSet<string> VariableSingleValues { get; } = new HashSet<string>();

        private static bool ContainsBooleans => typeof(T) == typeof(bool);

        private static bool ContainsIntegralNumbers => typeof(T) == typeof(long) 
                                                       || typeof(T) == typeof(int) 
                                                       || typeof(T) == typeof(short) 
                                                       || typeof(T) == typeof(byte);

        public string TypeName { get; set; }

        public bool FiltersAllValues
        {
            get
            {
                if (ContainsBooleans && CoversTrueFalse())
                {
                    return true;
                }

                var coversAll = false;
                var hasLessThanFilter = TryGetIsLessThanValue(out T lessThanValue);
                var hasGreaterThanFilter = TryGetIsGreaterThanValue(out T greaterThanValue);

                if (hasLessThanFilter && hasGreaterThanFilter)
                {
                    coversAll = lessThanValue.CompareTo(greaterThanValue) > 0
                        || lessThanValue.CompareTo(greaterThanValue) == 0 && SingleValues.Contains(lessThanValue)
                        || RangeValues.Any(rangeValue => rangeValue.Start.CompareTo(lessThanValue) <= 0 && rangeValue.End.CompareTo(greaterThanValue) >= 0);
                }

                if (ContainsIntegralNumbers && hasLessThanFilter && hasGreaterThanFilter && !coversAll)
                {
                    var lessThanIntegralNumber = ToLong(lessThanValue);
                    var greaterThanIntegralNumber = ToLong(greaterThanValue);
                    coversAll = greaterThanIntegralNumber - lessThanIntegralNumber + 1 <= RangesValuesCount() + SingleValues.Count
                        || greaterThanIntegralNumber == lessThanIntegralNumber && RangesFilterValue(lessThanValue);
                }
                return coversAll;
            }
        }

        public bool ContainsFilters => RangeValues.Any() 
                                       || VariableRanges.Any()
                                       || SingleValues.Any() || VariableSingleValues.Any()
                                       || RelationalOperators.Any()
                                       || TryGetIsLessThanValue(out T isLessThanValue) && isLessThanValue.CompareTo(_minExtent) != 0
                                       || TryGetIsGreaterThanValue(out T isGreaterThanValue) && isGreaterThanValue.CompareTo(_maxExtent) != 0;

        public void Add(IRangeClauseFilter filter)
        {
            var newFilter = (RangeClauseFilter<T>)filter;

            if (newFilter.TryGetIsLessThanValue(out T isLessThanValue))
            {
                AddIsClauseImpl(isLessThanValue, LogicSymbols.LT);
            }

            if (newFilter.TryGetIsGreaterThanValue(out T isGreaterThanValue))
            {
                AddIsClauseImpl(isGreaterThanValue, LogicSymbols.GT);
            }

            foreach (var range in newFilter.RangeValues)
            {
                AddValueRangeImpl(range.Start, range.End);
            }

            foreach (var range in newFilter.VariableRanges)
            {
                VariableRanges.Add(range);
            }

            foreach (var relationalOperator in newFilter.RelationalOperators)
            {
                AddRelationalOperatorImpl(relationalOperator);
            }

            foreach (var value in newFilter.SingleValues)
            {
                AddSingleValueImpl(value);
            }

            foreach (var value in newFilter.VariableSingleValues)
            {
                VariableSingleValues.Add(value);
            }
            _descriptorIsDirty = true;
        }

        public void AddIsClause(IParseTreeValue value, string operatorSymbol)
        {
            if (value.ParsesToConstantValue)
            {
                if (!_valueConverter(value, out T result))
                {
                    throw new ArgumentException();
                }
                AddIsClauseImpl(result, operatorSymbol);
            }
            else
            {
                AddRelationalOperatorImpl(value.ValueText);
            }
            _descriptorIsDirty = true;
        }

        public void AddRelationalOperator(IParseTreeValue value)
        {
            if (value.ParsesToConstantValue)
            {
                if (!_valueConverter(value, out T result))
                {
                    throw new ArgumentException(nameof(value));
                }
                AddSingleValueImpl(result);
            }
            else
            {
                AddRelationalOperatorImpl(value.ValueText);
            }
            _descriptorIsDirty = true;
        }

        public void AddSingleValue(IParseTreeValue value)
        {
            if (value.ParsesToConstantValue)
            {
                if (!_valueConverter(value, out T result))
                {
                    throw new ArgumentException();
                }
                AddSingleValueImpl(result);
            }
            else
            {
                VariableSingleValues.Add(value.ValueText);
            }
            _descriptorIsDirty = true;
        }

        public void AddValueRange(IParseTreeValue inputStartValue, IParseTreeValue inputEndValue)
        {
            var currentRanges = new List<(T Start, T End)>();
            currentRanges.AddRange(RangeValues);
            RangeValues.Clear();

            foreach (var range in currentRanges)
            {
                AddValueRangeImpl(range.Start, range.End);
            }

            if (inputStartValue.ParsesToConstantValue && inputEndValue.ParsesToConstantValue)
            {
                if (!(_valueConverter(inputStartValue, out T startValue) && _valueConverter(inputEndValue, out T endValue)))
                {
                    throw new ArgumentException();
                }
                AddValueRangeImpl(startValue, endValue);
            }
            else
            {
                AddVariableRangeImpl(inputStartValue.ValueText, inputEndValue.ValueText);
            }
            _descriptorIsDirty = true;
        }

        public IRangeClauseFilter FilterUnreachableClauses(IRangeClauseFilter filter)
        {
            if (filter is null)
            {
                throw new ArgumentNullException(nameof(filter));
            }

            if (!(filter is RangeClauseFilter<T>))
            {
                throw new ArgumentException($"Argument is not of type UCIRangeClauseFilter<{typeof(T)}>", "filter");
            }

            if (!ContainsFilters || filter.FiltersAllValues)
            {
                return _filterFactory.Create(TypeName, _valueFactory);
            }

            var filteredCoverage = (RangeClauseFilter<T>)MemberwiseClone();
            if (!filter.ContainsFilters && !_hasExtents)
            {
                return filteredCoverage;
            }

            filteredCoverage.RemoveClausesCoveredBy((RangeClauseFilter<T>)filter);

            return filteredCoverage;
        }

        public bool TryGetIsLessThanValue(out T isLessThanValue)
        {
            isLessThanValue = default;
            if (_isClause.TryGetValue(LogicSymbols.LT, out List<T> isLessThanValues) && isLessThanValues.Any())
            {
                isLessThanValue = isLessThanValues.Max();
                return true;
            }
            return false;
        }

        public bool TryGetIsGreaterThanValue(out T isGreaterThanValue)
        {
            isGreaterThanValue = default;
            if (_isClause.TryGetValue(LogicSymbols.GT, out List<T> isGreaterThanValues) && isGreaterThanValues.Any())
            {
                isGreaterThanValue = isGreaterThanValues.Min();
                return true;
            }
            return false;
        }

        public override bool Equals(object obj)
        {
            if (!(obj is RangeClauseFilter<T> filter))
            {
                return false;
            }
            if (filter.SingleValues.Count != SingleValues.Count
                || filter.VariableSingleValues.Count != VariableSingleValues.Count
                || filter.RangeValues.Count != RangeValues.Count
                || filter.VariableRanges.Count != VariableRanges.Count
                || filter.RelationalOperators.Count != RelationalOperators.Count)
            {
                return false;
            }

            if (filter.TryGetIsLessThanValue(out T isLessThanValue) 
                    && TryGetIsLessThanValue(out T myLessThanValue) 
                    && isLessThanValue.CompareTo(myLessThanValue) != 0
                || filter.TryGetIsGreaterThanValue(out T isGreaterThanValue) 
                    && TryGetIsGreaterThanValue(out T myGreaterThan) 
                    && isGreaterThanValue.CompareTo(myGreaterThan) != 0)
            {
                return false;
            }

            var hasSameRanges = filter.RangeValues.All(range => RangeValues.Contains(range))
                    && filter.VariableRanges.All(variableRange => VariableRanges.Contains(variableRange));

            var hasSameSingleValues = filter.SingleValues.All(value => SingleValues.Contains(value));
            var hasSameRelationalOperators = filter.RelationalOperators.All(relationalOperator => RelationalOperators.Contains(relationalOperator));
            return hasSameRanges && hasSameRelationalOperators && hasSameSingleValues;
        }

        public override string ToString()
        {
            if (!_descriptorIsDirty)
            {
                return _cachedDescriptor;
            }

            var descriptors = new HashSet<string>
            {
                GetIsClausesDescriptor(LogicSymbols.LT),
                GetIsClausesDescriptor(LogicSymbols.GT),
                GetRangesDescriptor(),
                GetSinglesDescriptor(),
                GetRelOpDescriptor()
            };

            descriptors.Remove(string.Empty);

            var descriptor = new StringBuilder();
            for (var idx = 0; idx < descriptors.Count; idx++)
            {
                if (idx > 0)
                {
                    descriptor.Append("!");
                }
                descriptor.Append(descriptors.ElementAt(idx));
            }
            _cachedDescriptor = descriptor.ToString();
            _descriptorIsDirty = false;
            return _cachedDescriptor;
        }

        public override int GetHashCode()
        {
            return ToString().GetHashCode();
        }

        public void AddExtents(IParseTreeValue min, IParseTreeValue max)
        {
            _hasExtents = true;
            if (_valueConverter(min, out _minExtent))
            {
                AddIsClauseImpl(_minExtent, LogicSymbols.LT);
            }

            if (_valueConverter(max, out _maxExtent))
            {
                AddIsClauseImpl(_maxExtent, LogicSymbols.GT);
            }
        }

        private bool FiltersAllRelationalOperators
        {
            get
            {
                if (ContainsBooleans)
                {
                    return CoversTrueFalse();
                }
                return SingleValues.Contains(_trueValue) && SingleValues.Contains(_falseValue)
                    || RangesFilterValue(_trueValue) && RangesFilterValue(_falseValue)
                    || IsLessThanFiltersValue(_trueValue) && IsLessThanFiltersValue(_falseValue)
                    || IsGreaterThanFiltersValue(_trueValue) && IsGreaterThanFiltersValue(_falseValue);
            }
        }

        private bool CoversTrueFalse()
        {
            return SingleValues.Contains(_trueValue) && SingleValues.Contains(_falseValue)
                || RangesFilterValue(_trueValue) && RangesFilterValue(_falseValue);
        }

        private void RemoveIsLessThanClause() => RemoveIsClauseImpl(LogicSymbols.LT);

        private void RemoveIsGreaterThanClause() => RemoveIsClauseImpl(LogicSymbols.GT);

        private void RemoveRangeValues(List<(T Start, T End)> toRemove)
        {
            foreach (var range in toRemove)
            {
                RangeValues.Remove(range);
            }
        }

        private void RemoveClausesCoveredBy(RangeClauseFilter<T> removalSpec)
        {
            RemoveIsClausesCoveredBy(removalSpec);
            RemoveRangesCoveredBy(removalSpec);
            RemoveSingleValuesCoveredBy(removalSpec);
            RemoveRelationalOperatorsCoveredBy(removalSpec);
        }

        private void RemoveIsClausesCoveredBy(RangeClauseFilter<T> removalSpec)
        {
            if (TryGetIsLessThanValue(out T isLessThanValue) 
                && removalSpec.TryGetIsLessThanValue(out T removalSpecLessThanValue)
                && removalSpecLessThanValue.CompareTo(isLessThanValue) >= 0)
            {
                RemoveIsLessThanClause();
            }

            if (TryGetIsGreaterThanValue(out T isGreaterThanValue)
                && removalSpec.TryGetIsGreaterThanValue(out T removalSpecGreaterThan)
                && removalSpecGreaterThan.CompareTo(isGreaterThanValue) <= 0)
            {
                RemoveIsGreaterThanClause();
            }
        }

        private void RemoveRangesCoveredBy(RangeClauseFilter<T> removalSpec)
        {
            if (!(RangeValues.Any() || VariableRanges.Any()))
            {
                return;
            }

            var rangesToRemove = new List<(T Start, T End)>();
            foreach (var range in RangeValues)
            {
                if (removalSpec.TryGetIsLessThanValue(out T removalSpecLessThanValue))
                {
                    if (removalSpecLessThanValue.CompareTo(range.Start) > 0
                        && removalSpecLessThanValue.CompareTo(range.End) > 0)
                    {
                        rangesToRemove.Add(range);
                        continue;
                    }
                }

                if (removalSpec.TryGetIsGreaterThanValue(out T removalSpecGreaterThanValue))
                {
                    if (removalSpecGreaterThanValue.CompareTo(range.Start) < 0
                        && removalSpecGreaterThanValue.CompareTo(range.End) < 0)
                    {
                        rangesToRemove.Add(range);
                        continue;
                    }
                }

                foreach (var removalRange in removalSpec.RangeValues)
                {
                    if (removalRange.Start.CompareTo(range.Start) <= 0 
                        && removalRange.End.CompareTo(range.End) >= 0)
                    {
                        rangesToRemove.Add(range);
                        break;
                    }
                }
            }
            RemoveRangeValues(rangesToRemove);

            var variableRangesToRemove = new List<string>();
            foreach(var variableRange in VariableRanges)
            {
                if (removalSpec.VariableRanges.Contains(variableRange))
                {
                    variableRangesToRemove.Add(variableRange);
                }
            }

            foreach (var variableRange in variableRangesToRemove)
            {
                VariableRanges.Remove(variableRange);
            }
        }

        private void RemoveSingleValuesCoveredBy(RangeClauseFilter<T> removalSpec)
        {
            List<T> toRemove = new List<T>();
            foreach (var singleValue in SingleValues)
            {
                if (removalSpec.TryGetIsLessThanValue(out T removalSpecLessThanValue))
                {
                    if (removalSpecLessThanValue.CompareTo(singleValue) > 0)
                    {
                        toRemove.Add(singleValue);
                        continue;
                    }
                }

                if (removalSpec.TryGetIsGreaterThanValue(out T removalSpecGreaterThanValue))
                {
                    if (removalSpecGreaterThanValue.CompareTo(singleValue) < 0)
                    {
                        toRemove.Add(singleValue);
                        continue;
                    }
                }

                foreach (var removalRange in removalSpec.RangeValues)
                {
                    if (removalRange.Item1.CompareTo(singleValue) <= 0
                        && removalRange.Item2.CompareTo(singleValue) >= 0)
                    {
                        toRemove.Add(singleValue);
                        break;
                    }
                }
            }
            toRemove.AddRange(removalSpec.SingleValues);

            foreach (var singleValue in toRemove)
            {
                SingleValues.Remove(singleValue);
            }

            var toRemoveVariables = new List<string>();
            foreach(var variable in VariableSingleValues)
            {
                if (removalSpec.VariableSingleValues.Contains(variable))
                {
                    toRemoveVariables.Add(variable);
                }
            }

            foreach (var variable in toRemoveVariables)
            {
                VariableSingleValues.Remove(variable);
            }
        }

        private void RemoveRelationalOperatorsCoveredBy(RangeClauseFilter<T> removalSpec)
        {
            List<string> toRemove = new List<string>();
            if (removalSpec.FiltersAllRelationalOperators)
            {
                RelationalOperators.Clear();
            }
            foreach (var removalOperators in removalSpec.RelationalOperators)
            {
                if (RelationalOperators.Contains(removalOperators))
                {
                    toRemove.Add(removalOperators);
                }
            }

            foreach (var relationalOperator in toRemove)
            {
                RelationalOperators.Remove(relationalOperator);
            }
        }

        private void AddIsClauseImpl(T val, string opSymbol)
        {
            if (ContainsBooleans)
            {
                AddIsClauseBoolean(val, opSymbol);
                return;
            }

            if (opSymbol.Equals(LogicSymbols.LT) || opSymbol.Equals(LogicSymbols.GT))
            {
                StoreIsClauseValue(val, opSymbol);
            }
            else if (opSymbol.Equals(LogicSymbols.LTE) || opSymbol.Equals(LogicSymbols.GTE))
            {
                var ltOrGtSymbol = opSymbol.Substring(0, opSymbol.Length - 1);
                StoreIsClauseValue(val, ltOrGtSymbol);

                AddSingleValueImpl(val);
            }
            else if (opSymbol.Equals(LogicSymbols.EQ))
            {
                AddSingleValueImpl(val);
            }
            else if (opSymbol.Equals(LogicSymbols.NEQ))
            {
                StoreIsClauseValue(val, LogicSymbols.LT);
                StoreIsClauseValue(val, LogicSymbols.GT);
            }

            FilterExistingRanges();
            FilterExistingSingles();
            TrimExistingRanges(true);
            TrimExistingRanges(false);
        }

        private void StoreIsClauseValue(T value, string opSymbol)
        {
            if (_isClause.ContainsKey(opSymbol))
            {
                _isClause[opSymbol].Add(value);
            }
            else
            {
                _isClause.Add(opSymbol, new List<T>() { value });
            }
        }

        private void AddIsClauseBoolean(T val, string opSymbol)
        {
            /*
             * Indeterminant cases are added as unresolved Relational Ops
             * 
            *************************** Is Clause Boolean Truth Table  *********************
            * 
            *                          Select Case Value
            *   Resolved Expression     True    False
            *   ****************************************************************************
            *   Is < True               False   False   <= Always False
            *   Is <= True              True    False   
            *   Is > True               False   True    
            *   Is >= True              True    True    <= Always True
            *   Is = True               True    False
            *   Is <> True              False   True
            *   Is > False              False   False   <= Always False
            *   Is >= False             False   True
            *   Is < False              True    False
            *   Is <= False             True    True    <= Always True
            *   Is = False              False   True
            *   Is <> False             True    False
            */

            var bVal = bool.Parse(val.ToString());

            if (opSymbol.Equals(LogicSymbols.NEQ)
                || opSymbol.Equals(LogicSymbols.EQ)
                || (opSymbol.Equals(LogicSymbols.GT) && bVal)
                || (opSymbol.Equals(LogicSymbols.LT) && !bVal)
                || (opSymbol.Equals(LogicSymbols.GTE) && !bVal)
                || (opSymbol.Equals(LogicSymbols.LTE) && bVal)
                )
            {
                AddRelationalOperatorImpl($"Is {opSymbol} {val}");
            }
            else if (opSymbol.Equals(LogicSymbols.GT) || opSymbol.Equals(LogicSymbols.GTE))
            {
                AddSingleValueImpl(ConvertToContainedGeneric(bVal));
            }
            else if (opSymbol.Equals(LogicSymbols.LT) || opSymbol.Equals(LogicSymbols.LTE))
            {
                AddSingleValueImpl(ConvertToContainedGeneric(!bVal));
            }
        }

        private void AddSingleValueImpl(T value)
        {
            if (IsClausesFilterValue(value)
                || RangesFilterValue(value))
            {
                return;
            }
            SingleValues.Add(value);
        }

        private void AddRelationalOperatorImpl(string value)
        {
            if (!FiltersAllRelationalOperators)
            {
                RelationalOperators.Add(value);
            }
        }

        private static long ToLong(T value)
        {
            return long.Parse(value.ToString());
        }

        private bool IsClausesFilterValue(T value) => IsLessThanFiltersValue(value) || IsGreaterThanFiltersValue(value);

        private bool IsLessThanFiltersValue(T value)
        {
            if (TryGetIsLessThanValue(out T isLT))
            {
                return value.CompareTo(isLT) < 0;
            }
            return false;
        }

        private bool IsGreaterThanFiltersValue(T value)
        {
            if (TryGetIsGreaterThanValue(out T isGT))
            {
                return value.CompareTo(isGT) > 0;
            }
            return false;
        }

        private void AddVariableRangeImpl(string inputStart, string inputEnd)
        {
            VariableRanges.Add($"{inputStart}:{inputEnd}");
        }

        private void AddValueRangeImpl(T inputStart, T inputEnd)
        {
            if (ContainsBooleans || inputStart.CompareTo(inputEnd) == 0)
            {
                SingleValues.Add(inputStart);
                SingleValues.Add(inputEnd);
                return;
            }

            var swapValueOrder = inputStart.CompareTo(inputEnd) > 0;
            T start = swapValueOrder ? inputEnd : inputStart;
            T end = swapValueOrder ? inputStart : inputEnd;

            if (IsClausesFilterRange(start, end) || RangesFilterRange(start, end))
            {
                return;
            }

            start = IsLessThanFiltersValue(start) ? _isClause[LogicSymbols.LT].Max() : start;
            end = IsGreaterThanFiltersValue(end) ? _isClause[LogicSymbols.GT].Min() : end;

            if (!RangeValues.Any())
            {
                RangeValues.Add((start, end));
            }
            else
            {
                var rangesToRemove = RangeValues.Where(rg => start.CompareTo(rg.Item1) <= 0 && end.CompareTo(rg.Item2) >= 0).ToList();
                rangesToRemove.ForEach(rtr => RangeValues.Remove(rtr));

                if (!TryMergeWithOverlappingRange(start, end))
                {
                    RangeValues.Add((start, end));
                }
            }

            ConcatenateExistingRanges();
            FilterExistingRanges();
            FilterExistingSingles();
        }

        private void ConcatenateExistingRanges()
        {
            if (!ContainsIntegralNumbers)
            {
                return;
            }

            if (RangeValues.Count() > 1)
            {
                int preConcatentateCount;
                do
                {
                    preConcatentateCount = RangeValues.Count();
                    ConcatenateRanges();
                } while (RangeValues.Count() < preConcatentateCount && RangeValues.Count() > 1);
            }
        }

        private void TrimExistingRanges(bool trimStart)
        {
            var rangesToTrim = (trimStart ? 
                 RangeValues.Where(rg => IsLessThanFiltersValue(rg.Start))
                 : RangeValues.Where(rg => IsGreaterThanFiltersValue(rg.End)))
                .ToList();

            var replacementRanges = new List<(T Start, T End)>();
            foreach (var range in rangesToTrim)
            {
                var newRange = trimStart 
                    ? (_isClause[LogicSymbols.LT].Max(), range.End)
                    : (range.Start, _isClause[LogicSymbols.GT].Min());

                replacementRanges.Add(newRange);
            }
            rangesToTrim.ForEach(rg => RangeValues.Remove(rg));
            RangeValues.AddRange(replacementRanges);
        }

        private void FilterExistingRanges()
            => RangeValues.Where(rg => IsClausesFilterRange(rg.Item1, rg.Item2))
            .ToList().ForEach(tr => RangeValues.Remove(tr));

        private void FilterExistingSingles()
            => SingleValues.Where(sv => IsClausesFilterValue(sv) || RangesFilterValue(sv))
            .ToList().ForEach(tr => SingleValues.Remove(tr));

        private bool IsClausesFilterRange(T start, T end)
            => IsLessThanFiltersValue(end) || IsGreaterThanFiltersValue(start);

        private bool RangesFilterRange(T start, T end)
            => RangeValues.Any(t => t.Item1.CompareTo(start) <= 0 && t.Item2.CompareTo(end) >= 0);

        private bool RangesFilterValue(T value) 
            => RangeValues.Any(rg => rg.Item1.CompareTo(value) <= 0 && rg.Item2.CompareTo(value) >= 0);

        private void ConcatenateRanges()
        {
            if (!ContainsIntegralNumbers)
            {
                return;
            }
            var concatenatedRanges = new List<(long Start, long End)>();
            var indexesToRemove = new List<int>();
            var sortedRanges = RangeValues.Select(range => ((long Start, long End))(ToLong(range.Start), ToLong(range.End)))
                                            .OrderBy(integralRange => integralRange.Start)
                                            .ToList();
            for (var idx = sortedRanges.Count - 1; idx > 0;)
            {
                if (sortedRanges[idx].Start - sortedRanges[idx - 1].End <= 1)
                {
                    concatenatedRanges.Add((sortedRanges[idx - 1].Start, sortedRanges[idx].End));
                    indexesToRemove.Add(idx);
                    indexesToRemove.Add(idx - 1);
                    break;
                }
                idx--;
            }
            //rebuild _ranges retaining the original order except placing the concatenated
            //range added to the end
            if (concatenatedRanges.Any())
            {
                var allRanges = new Dictionary<int, (T Start, T End)>();
                for (int idx = 0; idx < RangeValues.Count; idx++)
                {
                    allRanges.Add(idx, RangeValues[idx]);
                }

                indexesToRemove.ForEach(idx => sortedRanges.RemoveAt(idx));

                var tRanges = new List<(T Start, T End)>();
                foreach (var ral in sortedRanges)
                {
                    tRanges.Add((ConvertToContainedGeneric(ral.Start), ConvertToContainedGeneric(ral.End)));
                }

                foreach (var ral in concatenatedRanges)
                {
                    tRanges.Add((ConvertToContainedGeneric(ral.Start), ConvertToContainedGeneric(ral.End)));
                }

                RangeValues.Clear();
                foreach (var key in allRanges.Keys)
                {
                    if (tRanges.Contains(allRanges[key]))
                    {
                        RangeValues.Add(allRanges[key]);
                        tRanges.Remove(allRanges[key]);
                    }
                }
                RangeValues.AddRange(tRanges); //what's left is the concatenated result
            }
        }

        private int RangesValuesCount()
        {
            int result = 0;
            foreach (var range in RangeValues)
            {
                result = result + (int)(ToLong(range.Item2) - ToLong(range.Item1) + 1);
            }
            return result;
        }

        private T ConvertToContainedGeneric<K>(K value)
        {
            var parseTreeValue = _valueFactory.Create(value.ToString(), TypeName);
            if (_valueConverter(parseTreeValue, out T tValue))
            {
                return tValue;
            }
            throw new ArgumentException($"Unable to convert {value.ToString()} to {typeof(T).ToString()}");
        }

        private bool TryMergeWithOverlappingRange(T start, T end)
        {
            var endIsWithin = RangeValues.Where(t => t.Item1.CompareTo(end) < 0 && t.Item2.CompareTo(end) > 0);
            var startIsWithin = RangeValues.Where(t => t.Item1.CompareTo(start) < 0 && t.Item2.CompareTo(start) > 0);

            var rangeIsAdded = false;
            if (endIsWithin.Any() || startIsWithin.Any())
            {
                if (endIsWithin.Any())
                {
                    var original = endIsWithin.First();
                    RangeValues.Remove(endIsWithin.First());
                    RangeValues.Add((start, original.End));
                    rangeIsAdded = true;
                }
                else
                {
                    var original = startIsWithin.First();
                    RangeValues.Remove(startIsWithin.First());
                    RangeValues.Add((original.Start, end));
                    rangeIsAdded = true;
                }
            }

            return rangeIsAdded;
        }

        private void RemoveIsClauseImpl(string opSymbol)
        {
            if (_isClause.Keys.Contains(opSymbol))
            {
                if (_hasExtents)
                {
                    _isClause.Remove(opSymbol);
                    var extentVal = opSymbol.Equals(LogicSymbols.LT) ? _minExtent : _maxExtent;
                    AddIsClauseImpl(extentVal, opSymbol);
                }
                else
                {
                    _isClause.Remove(opSymbol);
                }
            }
        }

        private string GetSinglesDescriptor()
        {
            var singles = SingleValues.Select(sv => sv.ToString()).ToList();
            singles.AddRange(VariableSingleValues);
            return GetSingleValueTypeDescriptor(singles, "Single=");
        }

        private string GetRelOpDescriptor()
        {
            return GetSingleValueTypeDescriptor(RelationalOperators.ToList(), "RelOp=");
        }

        private string GetSingleValueTypeDescriptor<K>(List<K> values, string prefix)
        {
            if (!values.Any()){ return string.Empty; }

            StringBuilder series = new StringBuilder();
            foreach (var val in values)
            {
                series.Append($"{val},");
            }
            return $"{prefix}{series.ToString().Substring(0, series.Length - 1)}";
        }

        private string GetRangesDescriptor()
        {
            if (!(RangeValues.Any() || VariableRanges.Any())) { return string.Empty; }

            StringBuilder series = new StringBuilder();
            foreach (var val in RangeValues)
            {
                series.Append($"{val.Item1}:{val.Item2},");
            }
            foreach (var val in VariableRanges)
            {
                series.Append(val.ToString() + ",");
            }
            return $"Range={series.ToString().Substring(0, series.Length - 1)}";
        }

        private string GetIsClausesDescriptor(string opSymbol)
        {
            var result = string.Empty;
            if (_isClause.TryGetValue(opSymbol, out List<T> values))
            {
                var isLT = opSymbol.Equals(LogicSymbols.LT);
                var value = isLT ? values.Max() : values.Min();
                var extentToCompare = isLT ? _minExtent : _maxExtent;
                var prefix = isLT ? "IsLT=" : "IsGT=";
                if (!(_hasExtents && value.CompareTo(extentToCompare) == 0))
                {
                    result = $"{prefix}{value.ToString()}";
                }
            }
            return result;
        }
    }
}
