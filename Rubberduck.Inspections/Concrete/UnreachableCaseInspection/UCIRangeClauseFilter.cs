using Rubberduck.Parsing;
using System;
using System.Collections.Generic;
using System.Linq;

namespace Rubberduck.Inspections.Concrete.UnreachableCaseInspection
{
    public interface IUCIRangeClauseFilter
    {
        bool HasCoverage { get; }
        bool FiltersAllValues { get; }
        string TypeName { set; get; }
        IUCIRangeClauseFilter FilterUnreachableClauses(IUCIRangeClauseFilter filter);
        void Add(IUCIRangeClauseFilter newSummary);
        void AddValueRange(IUCIValue startVal, IUCIValue endVal);
        void AddIsClause(IUCIValue value, string opSymbol);
        void AddSingleValue(IUCIValue value);
        void AddRelationalOp(IUCIValue value);
    }

    public interface IUCIRangeClauseFilterTestSupport<T>
    {
        bool TryGetIsLTValue(out T isLT);
        bool TryGetIsGTValue(out T isGT);
        HashSet<T> SingleValues { get; }
    }

    public class UCIRangeClauseFilter<T> : IUCIRangeClauseFilter, IUCIRangeClauseFilterTestSupport<T> where T : IComparable<T>
    {
        private readonly IUCIValueFactory _valueFactory;
        private readonly IUCIRangeClauseFilterFactory _filterFactory;
        private readonly Func<IUCIValue, T> _tConverter;
        private readonly T _trueValue;
        private readonly T _falseValue;

        private List<Tuple<T, T>> _ranges;
        private Dictionary<string, List<T>> _isClause;
        private HashSet<T> _singleValues;
        private HashSet<string> _relationalOps;

        private bool _hasExtents;
        private T _minExtent;
        private T _maxExtent;

        public UCIRangeClauseFilter(string typeName, IUCIValueFactory valueFactory, IUCIRangeClauseFilterFactory summaryFactory, Func<IUCIValue, T> tConverter)
        {
            _valueFactory = valueFactory;
            _filterFactory = summaryFactory;
            _tConverter = tConverter;

            _ranges = new List<Tuple<T, T>>();
            _singleValues = new HashSet<T>();
            _isClause = new Dictionary<string, List<T>>();
            _relationalOps = new HashSet<string>();
            _hasExtents = false;
            _trueValue = _tConverter(_valueFactory.Create("True", typeName));
            _falseValue = _tConverter(_valueFactory.Create("False", typeName));
            TypeName = typeName;
        }

        //IUCIRangeClauseFilter
        public bool FiltersAllValues
        {
            get
            {

                var coversAll = false;
                var hasLTFilter = TryGetIsLTValue(out T ltValue);
                var hasGTFilter = TryGetIsGTValue(out T gtValue);
                if (hasLTFilter && hasGTFilter)
                {
                    coversAll = ltValue.CompareTo(gtValue) > 0
                        || ltValue.CompareTo(gtValue) == 0 && (SingleValues.Contains(ltValue))
                        || RangeValues.Any(rv => rv.Item1.CompareTo(ltValue) <= 0 && rv.Item2.CompareTo(gtValue) >= 0);
                }

                if (ContainsIntegerNumbers && hasLTFilter && hasGTFilter && !coversAll)
                {
                    var lt = ToLong(ltValue);
                    var gt = ToLong(gtValue);
                    coversAll = gt - lt + 1 <= RangesValuesCount() + SingleValues.Count()
                        || gt == lt && RangesFilterValue(ltValue);
                }
                else if (ContainsBooleans && !coversAll)
                {
                    coversAll = SingleValues.Contains(_trueValue);
                }
                return coversAll;
            }
        }

        //IUCIRangeClauseFilter
        public string TypeName { get; set; }

        //IUCIRangeClauseFilter
        public bool HasCoverage
        {
            get
            {
                return _ranges.Any()
                    || _singleValues.Any()
                    || TryGetIsLTValue(out T isLT) && isLT.CompareTo(_minExtent) != 0
                    || TryGetIsGTValue(out T isGT) && isGT.CompareTo(_maxExtent) != 0
                    || _relationalOps.Any();
            }
        }

        //IUCIRangeClauseFilter
        public void Add(IUCIRangeClauseFilter newSummary)
        {
            var itf = (UCIRangeClauseFilter<T>)newSummary;
            if (itf.TryGetIsLTValue(out T isLT))
            {
                AddIsClauseImpl(isLT, CompareTokens.LT);
            }
            if (itf.TryGetIsGTValue(out T isGT))
            {
                AddIsClauseImpl(isGT, CompareTokens.GT);
            }

            var ranges = itf.RangeValues;
            foreach (var tuple in ranges)
            {
                AddValueRangeImpl(tuple.Item1, tuple.Item2);
            }

            var relOps = itf.RelationalOps;
            foreach (var op in relOps)
            {
                AddRelationalOpImpl(op);
            }

            var singleVals = itf.SingleValues;
            foreach (var val in singleVals)
            {
                AddSingleValueImpl(val);
            }
        }

        //IUCIRangeClauseFilter
        public void AddIsClause(IUCIValue value, string opSymbol)
        {
            AddIsClauseImpl(_tConverter(value), opSymbol);
        }

        //IUCIRangeClauseFilter
        public void AddRelationalOp(IUCIValue value)
        {
            if (value.ParsesToConstantValue)
            {
                AddSingleValueImpl(_tConverter(value));
            }
            else
            {
                AddRelationalOpImpl(value.ValueText);
            }
        }

        //IUCIRangeClauseFilter
        public void AddSingleValue(IUCIValue value)
        {
            AddSingleValueImpl(_tConverter(value));
        }

        //IUCIRangeClauseFilter
        public void AddValueRange(IUCIValue inputStartVal, IUCIValue inputEndVal)
        {
            var currentRanges = new List<Tuple<T, T>>();
            currentRanges.AddRange(_ranges);
            _ranges.Clear();

            foreach (var range in currentRanges)
            {
                AddValueRangeImpl(range.Item1, range.Item2);
            }

            AddValueRangeImpl(_tConverter(inputStartVal), _tConverter(inputEndVal));
        }

        //IUCIRangeClauseFilter
        public IUCIRangeClauseFilter FilterUnreachableClauses(IUCIRangeClauseFilter filter)
        {
            if (filter is null)
            {
                throw new ArgumentNullException();
            }

            if (!(filter is UCIRangeClauseFilter<T>))
            {
                throw new ArgumentException($"Argument is not of type SummaryCoverage<{typeof(T).ToString()}>", "summary");
            }

            var filteredCoverage = _filterFactory.Create(TypeName, _valueFactory);

            filteredCoverage = (UCIRangeClauseFilter<T>)MemberwiseClone();
            if (!HasCoverage || filter.FiltersAllValues)
            {
                return _filterFactory.Create(TypeName, _valueFactory);
            }

            if (!filter.HasCoverage && !_hasExtents)
            {
                return filteredCoverage;
            }

            filteredCoverage = RemoveClausesCoveredBy((UCIRangeClauseFilter<T>)filteredCoverage, (UCIRangeClauseFilter<T>)filter);
            return filteredCoverage;
        }

        //IUCIRangeClauseFilterTestSupport<T>
        public bool TryGetIsLTValue(out T isLT)
        {
            isLT = default;
            if (_isClause.TryGetValue(CompareTokens.LT, out List<T> isLTValue))
            {
                if (isLTValue.Any())
                {
                    isLT = isLTValue.Max();
                    return true;
                }
            }
            return false;
        }

        //IUCIRangeClauseFilterTestSupport<T>
        public bool TryGetIsGTValue(out T isGT)
        {
            isGT = default;
            if (_isClause.TryGetValue(CompareTokens.GT, out List<T> isGTValue))
            {
                if (isGTValue.Any())
                {
                    isGT = isGTValue.Min();
                    return true;
                }
            }
            return false;
        }

        //IUCIRangeClauseFilterTestSupport<T>
        public HashSet<T> SingleValues => _singleValues;

        public override bool Equals(object obj)
        {
            if (!(obj is UCIRangeClauseFilter<T> filter))
            {
                return false;
            }
            if (!(filter is IUCIRangeClauseFilter))
            {
                return false;
            }

            if (filter.SingleValues.Count != SingleValues.Count
                || filter.RangeValues.Count != RangeValues.Count
                || filter.RelationalOps.Count != RelationalOps.Count)
            {
                return false;
            }

            var clausesMatch = true;
            if (filter.TryGetIsLTValue(out T isLT))
            {
                if (TryGetIsLTValue(out T myLT))
                {
                    if (isLT.CompareTo(myLT) != 0)
                    {
                        return false;
                    }
                }
                else
                {
                    return false;
                }
            }
            if (filter.TryGetIsGTValue(out T testGT))
            {
                if (TryGetIsGTValue(out T myGT))
                {
                    if (testGT.CompareTo(myGT) != 0)
                    {
                        return false;
                    }
                }
                else
                {
                    return false;
                }
            }
            if (!clausesMatch)
            {
                return false;
            }
            var theRanges = filter._ranges.All(rg => _ranges.Contains(rg));
            var singles = filter._singleValues.All(rg => _singleValues.Contains(rg));
            var relOps = filter._relationalOps.All(ro => _relationalOps.Contains(ro));
            return theRanges && relOps && singles;
        }

        public override string ToString()
        {
            var descriptors = new List<string>();
            descriptors = AddDescriptor(GetIsClausesDescriptor(CompareTokens.LT), descriptors);
            descriptors = AddDescriptor(GetIsClausesDescriptor(CompareTokens.GT), descriptors);
            descriptors = AddDescriptor(GetRangesDescriptor(), descriptors);
            descriptors = AddDescriptor(GetSinglesDescriptor(), descriptors);
            descriptors = AddDescriptor(GetRelOpDescriptor(), descriptors);
            var descriptor = string.Empty;
            foreach (var desc in descriptors)
            {
                descriptor = descriptor + desc + "!";
            }
            if (descriptor.Length > 0)
            {
                return descriptor.Substring(0, descriptor.Length - 1);
            }
            return string.Empty;
        }

        private static List<string> AddDescriptor(string descriptor, List<string> content)
        {
            if (descriptor.Length > 0)
            {
                content.Add(descriptor);
            }
            return content;
        }

        public override int GetHashCode()
        {
            return ToString().GetHashCode();
        }

        internal void AddExtents(IUCIValue min, IUCIValue max)
        {
            _hasExtents = true;
            _minExtent = _tConverter(min);
            _maxExtent = _tConverter(max);
            AddIsClauseImpl(_minExtent, CompareTokens.LT);
            AddIsClauseImpl(_maxExtent, CompareTokens.GT);
        }

        private bool FiltersAllRelationalOps
        {
            get
            {
                if (ContainsBooleans)
                {
                    return _singleValues.Contains(_trueValue)
                        || RangesFilterValue(_trueValue);
                }
                return _singleValues.Contains(_trueValue) && _singleValues.Contains(_falseValue)
                    || RangesFilterValue(_trueValue) && RangesFilterValue(_falseValue)
                    || IsLTFiltersValue(_trueValue) && IsLTFiltersValue(_falseValue)
                    || IsGTFiltersValue(_trueValue) && IsGTFiltersValue(_falseValue);
            }
        }

        private void RemoveIsLTClause()
        {
            RemoveIsClauseImpl(CompareTokens.LT);
        }

        private void RemoveIsGTClause()
        {
            RemoveIsClauseImpl(CompareTokens.GT);
        }

        private List<Tuple<T, T>> RangeValues => _ranges;

        private void RemoveRangeValues(List<Tuple<T, T>> toRemove)
        {
            foreach (var tp in toRemove)
            {
                _ranges.Remove(tp);
            }
        }

        private HashSet<string> RelationalOps => _relationalOps;

        private static bool ContainsBooleans => typeof(T) == typeof(bool);

        private static bool ContainsIntegerNumbers => typeof(T) == typeof(long) || typeof(T) == typeof(Int32) || typeof(T) == typeof(byte);

        private IUCIRangeClauseFilter RemoveClausesCoveredBy(UCIRangeClauseFilter<T> removeFrom, UCIRangeClauseFilter<T> removalSpec)
        {
            var newSummary = RemoveIsClausesCoveredBy(removeFrom, removalSpec);
            newSummary = RemoveRangesCoveredBy(removeFrom, removalSpec);
            newSummary = RemoveSingleValuesCoveredBy(removeFrom, removalSpec);
            newSummary = RemoveRelationalOpsCoveredBy(removeFrom, removalSpec);
            return newSummary;
        }

        private static UCIRangeClauseFilter<T> RemoveIsClausesCoveredBy(UCIRangeClauseFilter<T> removeFrom, UCIRangeClauseFilter<T> removalSpec)
        {
            if (removeFrom.TryGetIsLTValue(out T isLT))
            {
                if (removalSpec.TryGetIsLTValue(out T removalSpecLT))
                {
                    if (removalSpecLT.CompareTo(isLT) >= 0)
                    {
                        removeFrom.RemoveIsLTClause();
                    }
                }
            }
            if (removeFrom.TryGetIsGTValue(out T isGT))
            {
                if (removalSpec.TryGetIsGTValue(out T removalSpecGT))
                {
                    if (removalSpecGT.CompareTo(isGT) <= 0)
                    {
                        removeFrom.RemoveIsGTClause();
                    }
                }
            }
            return removeFrom;
        }

        private UCIRangeClauseFilter<T> RemoveRangesCoveredBy(UCIRangeClauseFilter<T> removeFrom, UCIRangeClauseFilter<T> removalSpec)
        {
            if (!(removeFrom.RangeValues.Any()))
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
            return removeFrom;
        }

        private UCIRangeClauseFilter<T> RemoveSingleValuesCoveredBy(UCIRangeClauseFilter<T> removeFrom, UCIRangeClauseFilter<T> removalSpec)
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
            return removeFrom;
        }

        private UCIRangeClauseFilter<T> RemoveRelationalOpsCoveredBy(UCIRangeClauseFilter<T> removeFrom, UCIRangeClauseFilter<T> removalSpec)
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

            if (opSymbol.Equals(CompareTokens.LT) || opSymbol.Equals(CompareTokens.GT))
            {
                if (!_isClause.Keys.Contains(opSymbol))
                {
                    _isClause.Add(opSymbol, new List<T>());
                }
                _isClause[opSymbol].Add(val);
            }
            else if (opSymbol.Equals(CompareTokens.LTE) || opSymbol.Equals(CompareTokens.GTE))
            {
                var ltOrGtSymbol = opSymbol.Substring(0, opSymbol.Length - 1);
                if (!_isClause.Keys.Contains(ltOrGtSymbol))
                {
                    _isClause.Add(ltOrGtSymbol, new List<T>());
                }
                _isClause[ltOrGtSymbol].Add(val);
                AddSingleValueImpl(val);
            }
            else if (opSymbol.Equals(CompareTokens.EQ))
            {
                AddSingleValueImpl(val);
            }
            else if (opSymbol.Equals(CompareTokens.NEQ))
            {
                _isClause[CompareTokens.LT].Add(val);
                _isClause[CompareTokens.GT].Add(val);
            }

            FilterExistingRanges();
            FilterExistingSingles();
            TrimExistingRanges(true);
            TrimExistingRanges(false);
        }

        private void AddIsClauseBoolean(T val, string opSymbol)
        {
            /*
             * Indeterminant cases are added as unresolved Relational Ops
             * 
            *************************** Is Clause Boolean Truth Table  *********************
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
            var relationalOpInput = $"Is {opSymbol} {val}";
            var bVal = bool.Parse(val.ToString());

            if(opSymbol.Equals(CompareTokens.NEQ)
                || opSymbol.Equals(CompareTokens.EQ)
                || (opSymbol.Equals(CompareTokens.GT) && bVal)
                || (opSymbol.Equals(CompareTokens.LT) && !bVal)
                || (opSymbol.Equals(CompareTokens.GTE) && !bVal)
                || (opSymbol.Equals(CompareTokens.LTE) && bVal)
                )
            {
                AddRelationalOpImpl(relationalOpInput);
            }
            else if (opSymbol.Equals(CompareTokens.GT) || opSymbol.Equals(CompareTokens.GTE))
            {
                AddSingleValueImpl(ConvertToContainedGeneric(bVal));
            }
            else if (opSymbol.Equals(CompareTokens.LT) || opSymbol.Equals(CompareTokens.LTE))
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

        private void AddValueRangeImpl(T inputStart, T inputEnd)
        {
            if (ContainsBooleans)
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

            start = IsLTFiltersValue(start) ? _isClause[CompareTokens.LT].Max() : start;
            end = IsGTFiltersValue(end) ? _isClause[CompareTokens.GT].Min() : end;

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
            if (!ContainsIntegerNumbers)
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
                    new Tuple<T, T>(_isClause[CompareTokens.LT].Max(), range.Item2)
                        : new Tuple<T, T>(range.Item1, _isClause[CompareTokens.GT].Min());

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
            if (!ContainsIntegerNumbers)
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
            var uciVal = _valueFactory.Create(value.ToString(), TypeName);
            return _tConverter(uciVal);
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
                    var extentVal = opSymbol.Equals(CompareTokens.LT) ? _minExtent : _maxExtent;
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
            return GetSingleValueTypeDescriptor(_singleValues, "Single=");
        }

        private string GetRelOpDescriptor()
        {
            return GetSingleValueTypeDescriptor(_relationalOps, "RelOp=");
        }

        private string GetSingleValueTypeDescriptor<K>(HashSet<K> values, string prefix)
        {
            var series = string.Empty;
            if (values.Any())
            {
                foreach (var val in values)
                {
                    series = series + val.ToString() + ",";
                }
                return $"{prefix}{series.Substring(0, series.Length - 1)}";
            }
            return series;
        }

        private string GetRangesDescriptor()
        {
            var series = string.Empty;
            if (_ranges.Any())
            {
                foreach (var val in _ranges)
                {
                    series = series + val.Item1.ToString() + ":" + val.Item2.ToString() + ",";
                }
                return $"Range={series.Substring(0, series.Length - 1)}";
            }
            return series;
        }

        private string GetIsClausesDescriptor(string opSymbol)
        {
            var result = string.Empty;
            if (_isClause.TryGetValue(opSymbol, out List<T> values))
            {
                var isLT = opSymbol.Equals(CompareTokens.LT);
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
