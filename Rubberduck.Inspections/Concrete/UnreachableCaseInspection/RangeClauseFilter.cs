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
        void AddValueRange(IParseTreeValue startVal, IParseTreeValue endVal);
        void AddIsClause(IParseTreeValue value, string opSymbol);
        void AddSingleValue(IParseTreeValue value);
        void AddRelationalOp(IParseTreeValue value);
    }

    public interface IRangeClauseFilterTestSupport<T>
    {
        bool TryGetIsLTValue(out T isLT);
        bool TryGetIsGTValue(out T isGT);
        HashSet<T> SingleValues { get; }
    }

    public class RangeClauseFilter<T> : IRangeClauseFilter, IRangeClauseFilterTestSupport<T> where T : IComparable<T>
    {
        private readonly IParseTreeValueFactory _valueFactory;
        private readonly IRangeClauseFilterFactory _filterFactory;
        private readonly TryConvertParseTreeValue<T> _valueConverter;
        private readonly T _trueValue;
        private readonly T _falseValue;

        private readonly List<Tuple<T, T>> _ranges;
        private readonly Dictionary<string, List<T>> _isClause;
        private readonly HashSet<T> _singleValues;
        private readonly HashSet<string> _relationalOps;
        private readonly HashSet<string> _variableRanges;
        private readonly HashSet<string> _variableSingles;

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

            _ranges = new List<Tuple<T, T>>();
            _singleValues = new HashSet<T>();
            _isClause = new Dictionary<string, List<T>>();
            _relationalOps = new HashSet<string>();
            _variableRanges = new HashSet<string>();
            _variableSingles = new HashSet<string>();
            _hasExtents = false;
            _falseValue = ConvertToContainedGeneric(false);
            _trueValue = ConvertToContainedGeneric(true);
            TypeName = typeName;
            _cachedDescriptor = string.Empty;
            _descriptorIsDirty = true;
        }

        private List<Tuple<T, T>> RangeValues => _ranges;

        private HashSet<string> VariableRanges => _variableRanges;

        private HashSet<string> RelationalOps => _relationalOps;

        public HashSet<T> SingleValues => _singleValues;

        private HashSet<string> VariableSingleValues => _variableSingles;

        private static bool ContainsBooleans => typeof(T) == typeof(bool);

        private static bool ContainsIntegralNumbers => typeof(T) == typeof(long) || typeof(T) == typeof(Int32) || typeof(T) == typeof(byte);

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
                var hasLTFilter = TryGetIsLTValue(out T ltValue);
                var hasGTFilter = TryGetIsGTValue(out T gtValue);
                if (hasLTFilter && hasGTFilter)
                {
                    coversAll = ltValue.CompareTo(gtValue) > 0
                        || ltValue.CompareTo(gtValue) == 0 && (SingleValues.Contains(ltValue))
                        || RangeValues.Any(rv => rv.Item1.CompareTo(ltValue) <= 0 && rv.Item2.CompareTo(gtValue) >= 0);
                }

                if (ContainsIntegralNumbers && hasLTFilter && hasGTFilter && !coversAll)
                {
                    var lt = ToLong(ltValue);
                    var gt = ToLong(gtValue);
                    coversAll = gt - lt + 1 <= RangesValuesCount() + SingleValues.Count()
                        || gt == lt && RangesFilterValue(ltValue);
                }
                return coversAll;
            }
        }

        public bool ContainsFilters
        {
            get
            {
                return _ranges.Any() || _variableRanges.Any()
                    || _singleValues.Any() || _variableSingles.Any()
                    || _relationalOps.Any()
                    || TryGetIsLTValue(out T isLT) && isLT.CompareTo(_minExtent) != 0
                    || TryGetIsGTValue(out T isGT) && isGT.CompareTo(_maxExtent) != 0;
            }
        }

        public void Add(IRangeClauseFilter filter)
        {
            var newFilter = (RangeClauseFilter<T>)filter;
            if (newFilter.TryGetIsLTValue(out T isLT))
            {
                AddIsClauseImpl(isLT, LogicSymbols.LT);
            }
            if (newFilter.TryGetIsGTValue(out T isGT))
            {
                AddIsClauseImpl(isGT, LogicSymbols.GT);
            }

            foreach (var tuple in newFilter.RangeValues)
            {
                AddValueRangeImpl(tuple.Item1, tuple.Item2);
            }

            foreach (var val in newFilter.VariableRanges)
            {
                _variableRanges.Add(val);
            }

            foreach (var op in newFilter.RelationalOps)
            {
                AddRelationalOpImpl(op);
            }

            foreach (var val in newFilter.SingleValues)
            {
                AddSingleValueImpl(val);
            }

            foreach (var val in newFilter.VariableSingleValues)
            {
                VariableSingleValues.Add(val);
            }
            _descriptorIsDirty = true;
        }

        public void AddIsClause(IParseTreeValue value, string opSymbol)
        {
            if (value.ParsesToConstantValue)
            {
                if (!_valueConverter(value, out T result))
                {
                    throw new ArgumentException();
                }
                AddIsClauseImpl(result, opSymbol);
            }
            else
            {
                AddRelationalOpImpl(value.ValueText);
            }
            _descriptorIsDirty = true;
        }

        public void AddRelationalOp(IParseTreeValue value)
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
                AddRelationalOpImpl(value.ValueText);
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

        public void AddValueRange(IParseTreeValue inputStartVal, IParseTreeValue inputEndVal)
        {
            var currentRanges = new List<Tuple<T, T>>();
            currentRanges.AddRange(_ranges);
            _ranges.Clear();

            foreach (var range in currentRanges)
            {
                AddValueRangeImpl(range.Item1, range.Item2);
            }

            if (inputStartVal.ParsesToConstantValue && inputEndVal.ParsesToConstantValue)
            {
                if (!(_valueConverter(inputStartVal, out T startVal) && _valueConverter(inputEndVal, out T endVal)))
                {
                    throw new ArgumentException();
                }
                AddValueRangeImpl(startVal, endVal);
            }
            else
            {
                AddVariableRangeImpl(inputStartVal.ValueText, inputEndVal.ValueText);
            }
            _descriptorIsDirty = true;
        }

        public IRangeClauseFilter FilterUnreachableClauses(IRangeClauseFilter filter)
        {
            if (filter is null)
            {
                throw new ArgumentNullException();
            }

            if (!(filter is RangeClauseFilter<T>))
            {
                throw new ArgumentException($"Argument is not of type UCIRangeClauseFilter<{typeof(T).ToString()}>", "filter");
            }

            var filteredCoverage = _filterFactory.Create(TypeName, _valueFactory);

            filteredCoverage = (RangeClauseFilter<T>)MemberwiseClone();
            if (!ContainsFilters || filter.FiltersAllValues)
            {
                return _filterFactory.Create(TypeName, _valueFactory);
            }

            if (!filter.ContainsFilters && !_hasExtents)
            {
                return filteredCoverage;
            }

            filteredCoverage = RemoveClausesCoveredBy((RangeClauseFilter<T>)filteredCoverage, (RangeClauseFilter<T>)filter);
            return filteredCoverage;
        }

        public bool TryGetIsLTValue(out T isLT)
        {
            isLT = default;
            if (_isClause.TryGetValue(LogicSymbols.LT, out List<T> isLTValues) && isLTValues.Any())
            {
                isLT = isLTValues.Max();
                return true;
            }
            return false;
        }

        public bool TryGetIsGTValue(out T isGT)
        {
            isGT = default;
            if (_isClause.TryGetValue(LogicSymbols.GT, out List<T> isGTValues) && isGTValues.Any())
            {
                isGT = isGTValues.Min();
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
                || filter.RelationalOps.Count != RelationalOps.Count)
            {
                return false;
            }

            if (filter.TryGetIsLTValue(out T isLT) && TryGetIsLTValue(out T myLT) && isLT.CompareTo(myLT) != 0
                || filter.TryGetIsGTValue(out T isGT) && TryGetIsGTValue(out T myGT) && isGT.CompareTo(myGT) != 0)
            {
                return false;
            }

            var theRanges = filter._ranges.All(rg => _ranges.Contains(rg))
                    && filter._variableRanges.All(vrg => _variableRanges.Contains(vrg));

            var singles = filter._singleValues.All(rg => _singleValues.Contains(rg));
            var relOps = filter._relationalOps.All(ro => _relationalOps.Contains(ro));
            return theRanges && relOps && singles;
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

        private bool FiltersAllRelationalOps
        {
            get
            {
                if (ContainsBooleans)
                {
                    return CoversTrueFalse();
                }
                return _singleValues.Contains(_trueValue) && _singleValues.Contains(_falseValue)
                    || RangesFilterValue(_trueValue) && RangesFilterValue(_falseValue)
                    || IsLTFiltersValue(_trueValue) && IsLTFiltersValue(_falseValue)
                    || IsGTFiltersValue(_trueValue) && IsGTFiltersValue(_falseValue);
            }
        }

        private bool CoversTrueFalse()
        {
            return _singleValues.Contains(_trueValue) && _singleValues.Contains(_falseValue)
                || RangesFilterValue(_trueValue) && RangesFilterValue(_falseValue);
        }

        private void RemoveIsLTClause() => RemoveIsClauseImpl(LogicSymbols.LT);

        private void RemoveIsGTClause() => RemoveIsClauseImpl(LogicSymbols.GT);

        private void RemoveRangeValues(List<Tuple<T, T>> toRemove)
        {
            foreach (var tp in toRemove)
            {
                _ranges.Remove(tp);
            }
        }

        private IRangeClauseFilter RemoveClausesCoveredBy(RangeClauseFilter<T> removeFrom, RangeClauseFilter<T> removalSpec)
        {
            var newFilter = RemoveIsClausesCoveredBy(removeFrom, removalSpec);
            newFilter = RemoveRangesCoveredBy(removeFrom, removalSpec);
            newFilter = RemoveSingleValuesCoveredBy(removeFrom, removalSpec);
            newFilter = RemoveRelationalOpsCoveredBy(removeFrom, removalSpec);
            return newFilter;
        }

        private static RangeClauseFilter<T> RemoveIsClausesCoveredBy(RangeClauseFilter<T> removeFrom, RangeClauseFilter<T> removalSpec)
        {
            if (removeFrom.TryGetIsLTValue(out T isLT) 
                && removalSpec.TryGetIsLTValue(out T removalSpecLT)
                && removalSpecLT.CompareTo(isLT) >= 0)
            {
                removeFrom.RemoveIsLTClause();
            }

            if (removeFrom.TryGetIsGTValue(out T isGT)
                && removalSpec.TryGetIsGTValue(out T removalSpecGT)
                && removalSpecGT.CompareTo(isGT) <= 0)
            {
                removeFrom.RemoveIsGTClause();
            }

            return removeFrom;
        }

        private RangeClauseFilter<T> RemoveRangesCoveredBy(RangeClauseFilter<T> removeFrom, RangeClauseFilter<T> removalSpec)
        {
            if (!(removeFrom.RangeValues.Any() || removeFrom.VariableRanges.Any()))
            {
                return removeFrom;
            }

            var rangesToRemove = new List<Tuple<T, T>>();
            if (removalSpec.TryGetIsLTValue(out T removalSpecLT))
            {
                foreach (var tup in removeFrom.RangeValues)
                {
                    if (removalSpecLT.CompareTo(tup.Item1) > 0 && removalSpecLT.CompareTo(tup.Item2) > 0)
                    {
                        rangesToRemove.Add(tup);
                    }
                }
            }

            if (removalSpec.TryGetIsGTValue(out T removalSpecGT))
            {
                foreach (var tup in removeFrom.RangeValues)
                {
                    if (removalSpecGT.CompareTo(tup.Item1) < 0 && removalSpecGT.CompareTo(tup.Item2) < 0)
                    {
                        rangesToRemove.Add(tup);
                    }
                }
            }

            foreach (var tup in removeFrom.RangeValues)
            {
                foreach (var rem in removalSpec.RangeValues)
                {
                    if (rem.Item1.CompareTo(tup.Item1) <= 0 && rem.Item2.CompareTo(tup.Item2) >= 0)
                    {
                        rangesToRemove.Add(tup);
                    }
                }
            }
            removeFrom.RemoveRangeValues(rangesToRemove);

            var varRangesToRemove = new List<string>();
            foreach(var value in removeFrom.VariableRanges)
            {
                if (removalSpec.VariableRanges.Contains(value))
                {
                    varRangesToRemove.Add(value);
                }
            }

            varRangesToRemove.ForEach(vr => removeFrom._variableRanges.Remove(vr));
            return removeFrom;
        }

        private RangeClauseFilter<T> RemoveSingleValuesCoveredBy(RangeClauseFilter<T> removeFrom, RangeClauseFilter<T> removalSpec)
        {
            List<T> toRemove = new List<T>();
            if (removalSpec.TryGetIsLTValue(out T removalSpecLT))
            {
                foreach (var sv in removeFrom.SingleValues)
                {
                    if (removalSpecLT.CompareTo(sv) > 0)
                    {
                        toRemove.Add(sv);
                    }
                }
            }

            if (removalSpec.TryGetIsGTValue(out T removalSpecGT))
            {
                foreach (var sv in removeFrom.SingleValues)
                {
                    if (removalSpecGT.CompareTo(sv) < 0)
                    {
                        toRemove.Add(sv);
                    }
                }
            }

            foreach (var tup in removalSpec.RangeValues)
            {
                foreach (var val in removeFrom.SingleValues)
                {
                    if (tup.Item1.CompareTo(val) <= 0 && tup.Item2.CompareTo(val) >= 0)
                    {
                        toRemove.Add(val);
                    }
                }
            }

            toRemove.AddRange(removalSpec.SingleValues);

            foreach (var rem in toRemove)
            {
                removeFrom.SingleValues.Remove(rem);
            }

            var toRemoveVariables = new List<string>();
            foreach(var value in removalSpec.VariableSingleValues)
            {
                if (removeFrom.VariableSingleValues.Contains(value))
                {
                    toRemoveVariables.Add(value);
                }
            }
            toRemoveVariables.ForEach(rv => removeFrom.VariableSingleValues.Remove(rv));
            return removeFrom;
        }

        private RangeClauseFilter<T> RemoveRelationalOpsCoveredBy(RangeClauseFilter<T> removeFrom, RangeClauseFilter<T> removalSpec)
        {
            List<string> toRemove = new List<string>();
            if (removalSpec.FiltersAllRelationalOps)
            {
                removeFrom.RelationalOps.Clear();
            }
            foreach (var rem in removalSpec.RelationalOps)
            {
                if (removeFrom.RelationalOps.Contains(rem))
                {
                    toRemove.Add(rem);
                }
            }

            foreach (var rem in toRemove)
            {
                removeFrom.RelationalOps.Remove(rem);
            }
            return removeFrom;
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
                AddRelationalOpImpl($"Is {opSymbol} {val}");
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
            _singleValues.Add(value);
        }

        private void AddRelationalOpImpl(string value)
        {
            if (!FiltersAllRelationalOps)
            {
                _relationalOps.Add(value);
                return;
            }
        }

        private static long ToLong(T value)
        {
            return long.Parse(value.ToString());
        }

        private bool IsClausesFilterValue(T value) => IsLTFiltersValue(value) || IsGTFiltersValue(value);

        private bool IsLTFiltersValue(T value)
        {
            if (TryGetIsLTValue(out T isLT))
            {
                return value.CompareTo(isLT) < 0;
            }
            return false;
        }

        private bool IsGTFiltersValue(T value)
        {
            if (TryGetIsGTValue(out T isGT))
            {
                return value.CompareTo(isGT) > 0;
            }
            return false;
        }

        private void AddVariableRangeImpl(string inputStart, string inputEnd)
        {
            _variableRanges.Add($"{inputStart}:{inputEnd}");
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

            start = IsLTFiltersValue(start) ? _isClause[LogicSymbols.LT].Max() : start;
            end = IsGTFiltersValue(end) ? _isClause[LogicSymbols.GT].Min() : end;

            if (!_ranges.Any())
            {
                var range = new Tuple<T, T>(start, end);
                _ranges.Add(range);
            }
            else
            {
                var rangesToRemove = _ranges.Where(rg => start.CompareTo(rg.Item1) <= 0 && end.CompareTo(rg.Item2) >= 0).ToList();
                rangesToRemove.ForEach(rtr => _ranges.Remove(rtr));

                if (!TryMergeWithOverlappingRange(start, end))
                {
                    _ranges.Add(new Tuple<T, T>(start, end));
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

            if (_ranges.Count() > 1)
            {
                int preConcatentateCount;
                do
                {
                    preConcatentateCount = _ranges.Count();
                    ConcatenateRanges();
                } while (_ranges.Count() < preConcatentateCount && _ranges.Count() > 1);
            }
        }

        private void TrimExistingRanges(bool trimStart)
        {
            var rangesToTrim = trimStart ? 
                 _ranges.Where(rg => IsLTFiltersValue(rg.Item1))
                 : _ranges.Where(rg => IsGTFiltersValue(rg.Item2));

            var replacementRanges = new List<Tuple<T, T>>();
            foreach (var range in rangesToTrim)
            {
                var newRange = trimStart ?
                    new Tuple<T, T>(_isClause[LogicSymbols.LT].Max(), range.Item2)
                        : new Tuple<T, T>(range.Item1, _isClause[LogicSymbols.GT].Min());

                replacementRanges.Add(newRange);
            }
            rangesToTrim.ToList().ForEach(rg => _ranges.Remove(rg));
            _ranges.AddRange(replacementRanges);
        }

        private void FilterExistingRanges()
            => _ranges.Where(rg => IsClausesFilterRange(rg.Item1, rg.Item2))
            .ToList().ForEach(tr => _ranges.Remove(tr));

        private void FilterExistingSingles()
            => _singleValues.Where(sv => IsClausesFilterValue(sv) || RangesFilterValue(sv))
            .ToList().ForEach(tr => _singleValues.Remove(tr));

        private bool IsClausesFilterRange(T start, T end)
            => IsLTFiltersValue(end) || IsGTFiltersValue(start);

        private bool RangesFilterRange(T start, T end)
            => _ranges.Any(t => t.Item1.CompareTo(start) <= 0 && t.Item2.CompareTo(end) >= 0);

        private bool RangesFilterValue(T value) 
            => _ranges.Any(rg => rg.Item1.CompareTo(value) <= 0 && rg.Item2.CompareTo(value) >= 0);

        private void ConcatenateRanges()
        {
            if (!ContainsIntegralNumbers)
            {
                return;
            }
            var concatenatedRanges = new List<Tuple<long, long>>();
            var indexesToRemove = new List<int>();
            var sortedRanges = _ranges.Select(rg => new Tuple<long, long>(ToLong(rg.Item1), ToLong(rg.Item2))).OrderBy(k => k.Item1).ToList();
            for (int idx = sortedRanges.Count() - 1; idx > 0;)
            {
                if (sortedRanges[idx].Item1 - sortedRanges[idx - 1].Item2 <= 1)
                {
                    concatenatedRanges.Add(new Tuple<long, long>(sortedRanges[idx - 1].Item1, sortedRanges[idx].Item2));
                    indexesToRemove.Add(idx);
                    indexesToRemove.Add(idx - 1);
                    idx = -1;
                }
                idx--;
            }
            //rebuild _ranges retaining the original order except placing the concatenated
            //range added to the end
            if (concatenatedRanges.Any())
            {
                var allRanges = new Dictionary<int, Tuple<T, T>>();
                for (int idx = 0; idx < _ranges.Count; idx++)
                {
                    allRanges.Add(idx, _ranges[idx]);
                }

                indexesToRemove.ForEach(idx => sortedRanges.RemoveAt(idx));

                var tRanges = new List<Tuple<T, T>>();
                foreach (var ral in sortedRanges)
                {
                    tRanges.Add(new Tuple<T, T>(ConvertToContainedGeneric(ral.Item1), ConvertToContainedGeneric(ral.Item2)));
                }

                foreach (var ral in concatenatedRanges)
                {
                    tRanges.Add(new Tuple<T, T>(ConvertToContainedGeneric(ral.Item1), ConvertToContainedGeneric(ral.Item2)));
                }

                _ranges.Clear();
                foreach (var key in allRanges.Keys)
                {
                    if (tRanges.Contains(allRanges[key]))
                    {
                        _ranges.Add(allRanges[key]);
                        tRanges.Remove(allRanges[key]);
                    }
                }
                _ranges.AddRange(tRanges); //what's left is the concatenated result
            }
        }

        private int RangesValuesCount()
        {
            int result = 0;
            foreach (var range in _ranges)
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
            var endIsWithin = _ranges.Where(t => t.Item1.CompareTo(end) < 0 && t.Item2.CompareTo(end) > 0);
            var startIsWithin = _ranges.Where(t => t.Item1.CompareTo(start) < 0 && t.Item2.CompareTo(start) > 0);

            var rangeIsAdded = false;
            if (endIsWithin.Any() || startIsWithin.Any())
            {
                if (endIsWithin.Any())
                {
                    var original = endIsWithin.First();
                    _ranges.Remove(endIsWithin.First());
                    _ranges.Add(new Tuple<T, T>(start, original.Item2));
                    rangeIsAdded = true;
                }
                else
                {
                    var original = startIsWithin.First();
                    _ranges.Remove(startIsWithin.First());
                    _ranges.Add(new Tuple<T, T>(original.Item1, end));
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
            var singles = _singleValues.Select(sv => sv.ToString()).ToList();
            singles.AddRange(_variableSingles);
            return GetSingleValueTypeDescriptor(singles, "Single=");
        }

        private string GetRelOpDescriptor()
        {
            return GetSingleValueTypeDescriptor(_relationalOps.ToList(), "RelOp=");
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
            if (!(_ranges.Any() || _variableRanges.Any())) { return string.Empty; }

            StringBuilder series = new StringBuilder();
            foreach (var val in _ranges)
            {
                series.Append($"{val.Item1}:{val.Item2},");
            }
            foreach (var val in _variableRanges)
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
