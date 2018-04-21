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
        void AddValueRange((IParseTreeValue StartValue, IParseTreeValue EndValue) valueRange);
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
                        || RangeValues.Any(range => range.Contains((lessThanValue, greaterThanValue)));
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
                AddValueRangeImpl(range);
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

        public void AddValueRange((IParseTreeValue StartValue, IParseTreeValue EndValue) valueRange)
        {
            var currentRanges = new List<(T Start, T End)>();
            currentRanges.AddRange(RangeValues);
            RangeValues.Clear();

            foreach (var range in currentRanges)
            {
                AddValueRangeImpl(range);
            }

            if (valueRange.StartValue.ParsesToConstantValue && valueRange.EndValue.ParsesToConstantValue)
            {
                AddValueRangeImpl(RangeFromValueRange(valueRange));
            }
            else
            {
                AddVariableRangeImpl(VariableRangeFromValueRange(valueRange));
            }
            _descriptorIsDirty = true;
        }

        private (T Start, T End) RangeFromValueRange((IParseTreeValue StartValue, IParseTreeValue EndValue) valueRange)
        {
            if (!(_valueConverter(valueRange.StartValue, out T startValue) && _valueConverter(valueRange.EndValue, out T endValue)))
            {
                throw new ArgumentException();
            }

            return (startValue, endValue);
        }

        private (string Start, string End) VariableRangeFromValueRange((IParseTreeValue StartValue, IParseTreeValue EndValue) valueRange)
        {
            return (valueRange.StartValue.ValueText, valueRange.EndValue.ValueText);
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
                GetSingleValuesDescriptor(),
                GetRelationalOperatorDescriptor()
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
                    if (removalSpecLessThanValue.IsGreaterThan(range))
                    {
                        rangesToRemove.Add(range);
                        continue;
                    }
                }

                if (removalSpec.TryGetIsGreaterThanValue(out T removalSpecGreaterThanValue))
                {
                    if (removalSpecGreaterThanValue.IsLessThan(range))
                    {
                        rangesToRemove.Add(range);
                        continue;
                    }
                }

                foreach (var removalRange in removalSpec.RangeValues)
                {
                    if (removalRange.Contains(range))
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
                    if (removalRange.Contains(singleValue))
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

        private void AddIsClauseImpl(T value, string operatorSymbol)
        {
            if (ContainsBooleans)
            {
                AddIsClauseBoolean(value, operatorSymbol);
                return;
            }

            if (operatorSymbol.Equals(LogicSymbols.LT) || operatorSymbol.Equals(LogicSymbols.GT))
            {
                StoreIsClauseValue(value, operatorSymbol);
            }
            else if (operatorSymbol.Equals(LogicSymbols.LTE) || operatorSymbol.Equals(LogicSymbols.GTE))
            {
                var lessThanOrGreaterThanSymbol = operatorSymbol.Substring(0, operatorSymbol.Length - 1);
                StoreIsClauseValue(value, lessThanOrGreaterThanSymbol);

                AddSingleValueImpl(value);
            }
            else if (operatorSymbol.Equals(LogicSymbols.EQ))
            {
                AddSingleValueImpl(value);
            }
            else if (operatorSymbol.Equals(LogicSymbols.NEQ))
            {
                StoreIsClauseValue(value, LogicSymbols.LT);
                StoreIsClauseValue(value, LogicSymbols.GT);
            }

            FilterExistingRanges();
            FilterExistingSingles();
            TrimExistingRanges(true);
            TrimExistingRanges(false);
        }

        private void StoreIsClauseValue(T value, string operatorSymbol)
        {
            if (_isClause.ContainsKey(operatorSymbol))
            {
                _isClause[operatorSymbol].Add(value);
            }
            else
            {
                _isClause.Add(operatorSymbol, new List<T>() { value });
            }
        }

        private void AddIsClauseBoolean(T value, string operatorSymbol)
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

            var booleanValue = bool.Parse(value.ToString());

            if (operatorSymbol.Equals(LogicSymbols.NEQ)
                || operatorSymbol.Equals(LogicSymbols.EQ)
                || (operatorSymbol.Equals(LogicSymbols.GT) && booleanValue)
                || (operatorSymbol.Equals(LogicSymbols.LT) && !booleanValue)
                || (operatorSymbol.Equals(LogicSymbols.GTE) && !booleanValue)
                || (operatorSymbol.Equals(LogicSymbols.LTE) && booleanValue)
                )
            {
                AddRelationalOperatorImpl($"Is {operatorSymbol} {value}");
            }
            else if (operatorSymbol.Equals(LogicSymbols.GT) || operatorSymbol.Equals(LogicSymbols.GTE))
            {
                AddSingleValueImpl(ConvertToContainedGeneric(booleanValue));
            }
            else if (operatorSymbol.Equals(LogicSymbols.LT) || operatorSymbol.Equals(LogicSymbols.LTE))
            {
                AddSingleValueImpl(ConvertToContainedGeneric(!booleanValue));
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
            if (TryGetIsLessThanValue(out T isLessThanValue))
            {
                return value.CompareTo(isLessThanValue) < 0;
            }
            return false;
        }

        private bool IsGreaterThanFiltersValue(T value)
        {
            if (TryGetIsGreaterThanValue(out T isGreaterThanValue))
            {
                return value.CompareTo(isGreaterThanValue) > 0;
            }
            return false;
        }

        private void AddVariableRangeImpl((string Start, string End) variableRange)
        {
            VariableRanges.Add($"{variableRange.Start}:{variableRange.End}");
        }

        private void AddValueRangeImpl((T Start, T End) range)
        {
            if (ContainsBooleans || range.Start.CompareTo(range.End) == 0)
            {
                SingleValues.Add(range.Start);
                SingleValues.Add(range.End);
                return;
            }

            var orderedRange = OrderedRange(range);

            if (IsClausesFilterRange(orderedRange) || RangesFilterRange(orderedRange))
            {
                return;
            }

            var extendedRange = RangeExtendedToIsClauseBoundaries(orderedRange);

            if (!RangeValues.Any())
            {
                RangeValues.Add(extendedRange);
            }
            else
            {
                var rangesToRemove = RangeValues.Where(storedRange => 
                                                    extendedRange.Start.CompareTo(storedRange.Start) <= 0 
                                                    && extendedRange.End.CompareTo(storedRange.End) >= 0)
                                                .ToList();
                rangesToRemove.ForEach(rangeToRemove => RangeValues.Remove(rangeToRemove));

                if (!TryMergeWithOverlappingRange(extendedRange))
                {
                    RangeValues.Add(extendedRange);
                }
            }

            ConcatenateExistingRanges();
            FilterExistingRanges();
            FilterExistingSingles();
        }

        private (T Start, T End) OrderedRange((T Start, T End) range)
        {
            if (range.Start.CompareTo(range.End) > 0)
            {
                return (range.End, range.Start);
            }

            return range;
        }

        private (T Start, T End) RangeExtendedToIsClauseBoundaries((T Start, T End) range)
        {
            var start = IsLessThanFiltersValue(range.Start) ? _isClause[LogicSymbols.LT].Max() : range.Start;
            var end = IsGreaterThanFiltersValue(range.End) ? _isClause[LogicSymbols.GT].Min() : range.End;

            return (start, end);
        }

        private void ConcatenateExistingRanges()
        {
            if (!ContainsIntegralNumbers)
            {
                return;
            }

            if (RangeValues.Count > 1)
            {
                int preConcatentateCount;
                do
                {
                    preConcatentateCount = RangeValues.Count;
                    ConcatenateRanges();
                } while (RangeValues.Count < preConcatentateCount && RangeValues.Count > 1);
            }
        }

        private void TrimExistingRanges(bool trimStart)
        {
            var rangesToTrim = (trimStart  
                 ? RangeValues.Where(rg => IsLessThanFiltersValue(rg.Start))
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
            => RangeValues.Where(range => IsClausesFilterRange(range))
                            .ToList()
                            .ForEach(filteredRange => RangeValues.Remove(filteredRange));

        private void FilterExistingSingles()
            => SingleValues.Where(singleValue => IsClausesFilterValue(singleValue) || RangesFilterValue(singleValue))
                            .ToList()
                            .ForEach(filteredSingleValue => SingleValues.Remove(filteredSingleValue));

        private bool IsClausesFilterRange((T Start, T End) range)
            => IsLessThanFiltersValue(range.End) || IsGreaterThanFiltersValue(range.Start);

        private bool RangesFilterRange((T Start, T End) range)
            => RangeValues.Any(storedRange => storedRange.Start.CompareTo(range.Start) <= 0 && storedRange.End.CompareTo(range.End) >= 0);

        private bool RangesFilterValue(T value) 
            => RangeValues.Any(range => range.Start.CompareTo(value) <= 0 && range.End.CompareTo(value) >= 0);

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
            throw new ArgumentException($"Unable to convert {value.ToString()} to {typeof(T)}");
        }

        private bool TryMergeWithOverlappingRange((T Start, T End) range)
        {
            if (RangeValues.Any(storedRange => storedRange.Contains(range)))
            {
                //Nothing to do here; merge with the containing range will result in the containing range.
                return true;
            }

            var originalRangeEnclosingEnd = RangeValues.Where(storedRange => storedRange.Encloses(range.End))
                                                        .Cast<(T Start, T End)?>()
                                                        .FirstOrDefault();
            if (originalRangeEnclosingEnd != null)
            {
                var original = originalRangeEnclosingEnd.Value;
                RangeValues.Remove(original);
                RangeValues.Add((range.Start, original.End));
                return true;
            }

            var originalRangeEnclosingStart = RangeValues.Where(storedRange => storedRange.Encloses(range.Start))
                                                            .Cast<(T Start, T End)?>()
                                                            .FirstOrDefault();
            if (originalRangeEnclosingStart != null)
            {
                var original = originalRangeEnclosingStart.Value;
                RangeValues.Remove(original);
                RangeValues.Add((original.Start, range.End));
                return true;
            }

            return false;
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

        private string GetSingleValuesDescriptor()
        {
            var singleValueTexts = SingleValues.Select(singleValue => singleValue.ToString()).ToList();
            singleValueTexts.AddRange(VariableSingleValues);
            return GetSingleValueTypeDescriptor(singleValueTexts, "Single=");
        }

        private string GetRelationalOperatorDescriptor()
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

    public static class ComparisonExtensions
    {
        public static bool IsContainedIn<T>(this T value, (T Start, T End) range) where T: IComparable<T>
        {
            return range.Start.CompareTo(value) <= 0 && range.End.CompareTo(value) >= 0;
        }

        public static bool IsContainedIn<T>(this (T Start, T End) range, (T Start, T End) otherRange) where T : IComparable<T>
        {
            return otherRange.Start.CompareTo(range.Start) <= 0 && otherRange.End.CompareTo(range.End) >= 0;
        }

        public static bool IsEnclosedBy<T>(this T value, (T Start, T End) range) where T : IComparable<T>
        {
            return range.Start.CompareTo(value) < 0 && range.End.CompareTo(value) > 0;
        }

        public static bool IsEnclosedBy<T>(this (T Start, T End) range, (T Start, T End) otherRange) where T : IComparable<T>
        {
            return otherRange.Start.CompareTo(range.Start) < 0 && otherRange.End.CompareTo(range.End) > 0;
        }


        public static bool Contains<T>(this (T Start, T End) range, T value) where T : IComparable<T>
        {
            return value.IsContainedIn(range);
        }

        public static bool Contains<T>(this (T Start, T End) range, (T Start, T End) otherRange) where T : IComparable<T>
        {
            return otherRange.IsContainedIn(range);
        }

        public static bool Encloses<T>(this (T Start, T End) range, T value) where T : IComparable<T>
        {
            return value.IsEnclosedBy(range);
        }

        public static bool Encloses<T>(this (T Start, T End) range, (T Start, T End) otherRange) where T : IComparable<T>
        {
            return otherRange.IsEnclosedBy(range);
        }


        public static bool IsLessThan<T>(this T value, (T Start, T End) range) where T : IComparable<T>
        {
            return range.Start.CompareTo(value) > 0 && range.End.CompareTo(value) > 0;
        }

        public static bool IsGreaterThan<T>(this T value, (T Start, T End) range) where T : IComparable<T>
        {
            return range.Start.CompareTo(value) < 0 && range.End.CompareTo(value) < 0;
        }
    }
}
