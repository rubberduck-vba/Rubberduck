using Rubberduck.Parsing;
using System;
using System.Collections.Generic;
using System.Linq;

namespace Rubberduck.Inspections.Concrete.UnreachableSelectCase
{
    internal static class CompareTokens
    {
        public static readonly string EQ = "=";
        public static readonly string NEQ = "<>";
        public static readonly string LT = "<";
        public static readonly string LTE = "<=";
        public static readonly string GT = ">";
        public static readonly string GTE = ">=";
    }

    public interface IUCIRangeClauseFilter
    {
        bool HasCoverage { get; }
        bool FiltersAllValues { get; }
        bool CoversTrueFalse { get; }
        string TypeName { set; get; }
        IUCIRangeClauseFilter FilterUnreachableClauses(IUCIRangeClauseFilter filter);
        void Add(IUCIRangeClauseFilter newSummary);
        void AddValueRange(IUCIValue startVal, IUCIValue endVal);
        void AddIsClause(IUCIValue value, string opSymbol);
        void AddSingleValue(IUCIValue value);
        void AddRelationalOp(IUCIValue value);
        void AddExtents(IUCIValue minValue, IUCIValue maxValue);
    }

    public interface IUCIRangeFilterTestSupport<T>
    {
        bool TryGetIsLTValue(out T isLT);
        void RemoveIsLTClause();
        bool TryGetIsGTValue(out T isGT);
        void RemoveIsGTClause();
        List<Tuple<T, T>> RangeValues { get; }
        void RemoveRangeValues(List<Tuple<T, T>> toRemove);
        HashSet<string> RelationalOps { get; }
        HashSet<T> SingleValues { get; }
    }

    public class UCIRangeClauseFilter<T> : IUCIRangeClauseFilter, IUCIRangeFilterTestSupport<T> where T : IComparable<T>
    {
        private readonly IUCIValueFactory _valueFactory;
        private readonly IUCIRangeClauseFilterFactory _summaryFactory;
        private readonly Func<IUCIValue, T> _tConverter;
        private readonly T _trueValue;
        private readonly T _falseValue;

        private List<Tuple<T, T>> _ranges;
        private Dictionary<string, List<T>> _isClause;
        private HashSet<T> _singleValues;
        private HashSet<string> _relationalOps;

        private T _minExtent;
        bool _hasExtents;
        private T _maxExtent;

        private static bool ContainsBooleans => typeof(T) == typeof(bool);
        private static bool ContainsIntegerNumbers => typeof(T) == typeof(long) || typeof(T) == typeof(Int32) || typeof(T) == typeof(byte);

        public UCIRangeClauseFilter(string typeName, IUCIValueFactory valueFactory, IUCIRangeClauseFilterFactory summaryFactory, Func<IUCIValue, T> tConverter)
        {
            _valueFactory = valueFactory;
            _summaryFactory = summaryFactory;
            _tConverter = tConverter;

            _ranges = new List<Tuple<T, T>>();
            _singleValues = new HashSet<T>();
            _isClause = new Dictionary<string, List<T>>();
            _relationalOps = new HashSet<string>();
            _hasExtents = false;
            _trueValue = _tConverter(_valueFactory.Create("True", TypeName));
            _falseValue = _tConverter(_valueFactory.Create("False", TypeName));
            TypeName = typeName;
        }

        //ISummaryCoverage
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
                        || gt == lt && RangesCoverValue(ltValue);
                }
                else if (ContainsBooleans && !coversAll)
                {
                    //TODO: Add test and code for IsLT = 1 and IsGT = -1
                    coversAll = SingleValues.Count == 2;
                }
                return coversAll;
            }
        }

        //ISummaryCoverage
        public bool CoversTrueFalse => _singleValues.Contains(_trueValue) && _singleValues.Contains(_falseValue)
            || _ranges.Any(rg => rg.Item1.CompareTo(_trueValue) <= 0 && rg.Item2.CompareTo(_falseValue) >= 0)
            || IsClausesCoversTrueFalse();

        //ISummaryCoverage
        public string TypeName { get; set; }

        //ISummaryCoverage
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

        //ISummaryCoverage
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
                AddRelationalOpImpl(op); // _valueFactory.Create(op,TypeName));
            }

            var singleVals = itf.SingleValues;
            foreach (var val in singleVals)
            {
                AddSingleValueImpl(val);
            }
        }

        //ISummaryCoverage
        public void AddExtents(IUCIValue min, IUCIValue max)
        {
            _hasExtents = true;
            _minExtent = _tConverter(min);
            _maxExtent = _tConverter(max);
            AddIsClauseImpl(_minExtent, CompareTokens.LT);
            AddIsClauseImpl(_maxExtent, CompareTokens.GT);
        }

        //ISummaryCoverage
        public void AddIsClause(IUCIValue value, string opSymbol)
        {
            AddIsClauseImpl(_tConverter(value), opSymbol);
        }

        //ISummaryCoverage
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

        //ISummaryCoverage
        public void AddSingleValue(IUCIValue value)
        {
            AddSingleValueImpl(_tConverter(value));
        }

        //ISummaryCoverage
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

        //ISummaryCoverage
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

            var filteredCoverage = _summaryFactory.Create(TypeName, _valueFactory, _summaryFactory);

            filteredCoverage = (UCIRangeClauseFilter<T>)MemberwiseClone();
            if (!HasCoverage || filter.FiltersAllValues)
            {
                return _summaryFactory.Create(TypeName, _valueFactory, _summaryFactory);
            }

            if (!filter.HasCoverage && !_hasExtents)
            {
                return filteredCoverage;
            }

            filteredCoverage = RemoveClausesCoveredBy((UCIRangeClauseFilter<T>)filteredCoverage, (UCIRangeClauseFilter<T>)filter);
            return filteredCoverage;
        }

        public override bool Equals(object obj)
        {
            if (!(obj is UCIRangeClauseFilter<T> element))
            {
                return false;
            }
            if (!(element is IUCIRangeClauseFilter))
            {
                return false;
            }
            var test = element;
            if (test.SingleValues.Count != SingleValues.Count
                || test.RangeValues.Count != RangeValues.Count
                || test.RelationalOps.Count != RelationalOps.Count)
            {
                return false;
            }

            var clausesMatch = true;
            if (test.TryGetIsLTValue(out T isLT))
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
            if (test.TryGetIsGTValue(out T testGT))
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
            var theRanges = test._ranges.All(rg => _ranges.Contains(rg));
            var singles = test._singleValues.All(rg => _singleValues.Contains(rg));
            var relOps = test._relationalOps.All(ro => _relationalOps.Contains(ro));
            return theRanges && relOps && singles;
        }

        public override string ToString()
        {
            var descriptors = new List<string>();
            descriptors = AddDesciptor(GetIsLTClausesDescriptor(), descriptors);
            descriptors = AddDesciptor(GetIsGTClausesDescriptor(), descriptors);
            descriptors = AddDesciptor(GetRangesDescriptor(), descriptors);
            descriptors = AddDesciptor(GetSinglesDescriptor(), descriptors);
            descriptors = AddDesciptor(GetRelOpDescriptor(), descriptors);
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

        public override int GetHashCode()
        {
            return ToString().GetHashCode();
        }

        //ISummaryCoverageTestSupport<T>
        public void RemoveIsLTClause()
        {
            RemoveIsClauseImpl(CompareTokens.LT);
        }

        //ISummaryCoverageTestSupport<T>
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

        //ISummaryCoverageTestSupport<T>
        public void RemoveIsGTClause()
        {
            RemoveIsClauseImpl(CompareTokens.GT);
        }

        //ISummaryCoverageTestSupport<T>
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

        //ISummaryCoverageTestSupport<T>
        public List<Tuple<T, T>> RangeValues => _ranges;

        //ISummaryCoverageTestSupport<T>
        public void RemoveRangeValues(List<Tuple<T, T>> toRemove)
        {
            foreach (var tp in toRemove)
            {
                _ranges.Remove(tp);
            }
        }

        //ISummaryCoverageTestSupport<T>
        public HashSet<string> RelationalOps => _relationalOps;

        //ISummaryCoverageTestSupport<T>
        public HashSet<T> SingleValues => _singleValues;

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
            if (((IUCIRangeClauseFilter)removalSpec).CoversTrueFalse)
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

        private HashSet<long> FlattenRangeValuesAsLongs()
        {
            var discreteRangeValues = new HashSet<long>();
            if (ContainsIntegerNumbers)
            {
                foreach (var range in _ranges)
                {
                    for (var value = ToLong(range.Item1); value <= ToLong(range.Item2); value++)
                    {
                        discreteRangeValues.Add(value);
                    }
                }
            }
            return discreteRangeValues;
        }

        private void AddIsClauseImpl(T val, string opSymbol)
        {
            if (ContainsBooleans)
            {
                var bVal = bool.Parse(val.ToString());
                //TODO: verify empirical test - do not have data for Is <> 1
                if (opSymbol.Equals(CompareTokens.NEQ))
                {
                    AddSingleValueImpl(ConvertToT(!bVal));
                }
                else if (opSymbol.Equals(CompareTokens.EQ))
                {
                    AddSingleValueImpl(ConvertToT(bVal));
                }
                else if (opSymbol.Equals(CompareTokens.GT))
                {
                    if (bVal)
                    {
                        AddSingleValueImpl(ConvertToT(!bVal));
                    }
                    //TODO: verify empirical test
                    //TODO: test that this results in a unreachable case
                    //Empirical testing indicates Is > false does not trigger on true or false
                }
                else if (opSymbol.Equals(CompareTokens.LT))
                {
                    if (!bVal)
                    {
                        AddSingleValueImpl(ConvertToT(!bVal));
                    }
                    //TODO: test that this results in a unreachable case
                    //TODO: verify empirical test
                    //Empirical testing indicates Is < true does not trigger on true or false
                }
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
        }

        private void AddSingleValueImpl(T value)
        {
            if (IsLTCoversValue(value) 
                || IsGTCoversValue(value)
                || RangesCoverValue(value))
            {
                return;
            }
            _singleValues.Add(value);
        }

        private void AddRelationalOpImpl(string value)
        {
            if (!CoversTrueFalse)
            {
                _relationalOps.Add(value);
                return;
            }
        }

        private static long ToLong(T value)
        {
            return long.Parse(value.ToString());
        }

        private bool IsLTCoversValue(T value)
        {
            if (TryGetIsLTValue(out T isLT))
            {
                return value.CompareTo(isLT) < 0;
            }
            return false;
        }

        private bool IsGTCoversValue(T value)
        {
            if (TryGetIsGTValue(out T isGT))
            {
                return value.CompareTo(isGT) > 0;
            }
            return false;
        }

        private void AddValueRangeImpl(T inputStart, T inputEnd)
        {
            var swapValueOrder = inputStart.CompareTo(inputEnd) > 0;
            T start = swapValueOrder ? inputEnd : inputStart;
            T end = swapValueOrder ? inputStart : inputEnd;

            if (ContainsBooleans)
            {
                SingleValues.Add(start);
                SingleValues.Add(end);
                return;
            }

            if (IsLTCoversValue(end) 
                || IsGTCoversValue(start)
                || _ranges.Any(t => t.Item1.CompareTo(start) <= 0 && t.Item2.CompareTo(end) >= 0))
            {
                return;
            }

            start = IsLTCoversValue(start) ? GetIsLTValue() : start;
            end = IsGTCoversValue(end) ? GetIsGTValue() : end;

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

            if (_ranges.Count() > 1)
            {
                int preConcatentateCount;
                do
                {
                    preConcatentateCount = _ranges.Count();
                    ConcatenateRanges();
                } while (_ranges.Count() < preConcatentateCount && _ranges.Count() > 1);
            }
            _ranges.ForEach(rg => RemoveSinglesCoveredByRange(rg));
        }

        private void RemoveSinglesCoveredByRange(Tuple<T,T> range)
        {
            var toRemove = _singleValues.Where(sv => range.Item1.CompareTo(sv) <= 0 && range.Item2.CompareTo(sv) >= 0).ToList();
            toRemove.ForEach(tr => _singleValues.Remove(tr));
        }

        private bool CoversRange(T start, T end)
        {
            var existingCoversProposed = _ranges.Where(t => t.Item1.CompareTo(start) <= 0 && t.Item2.CompareTo(end) >= 0);
            return existingCoversProposed.Any();
        }

        private bool RangesCoverValue(T value)
        {
            return _ranges.Any(rg => rg.Item1.CompareTo(value) <= 0 && rg.Item2.CompareTo(value) >= 0);
        }

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
                for(int idx = 0; idx < _ranges.Count; idx++)
                {
                    allRanges.Add(idx, _ranges[idx]);
                }

                indexesToRemove.ForEach(idx => sortedRanges.RemoveAt(idx));

                var tRanges = new List<Tuple<T, T>>();
                foreach (var ral in sortedRanges)
                {
                    tRanges.Add( new Tuple<T, T>(ConvertToT(ral.Item1), ConvertToT(ral.Item2)));
                }

                foreach (var ral in concatenatedRanges)
                {
                    tRanges.Add(new Tuple<T, T>(ConvertToT(ral.Item1), ConvertToT(ral.Item2)));
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

        private T ConvertToT<K>(K value)
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

        private static List<string> AddDesciptor(string descriptor, List<string> content)
        {
            if (descriptor.Length > 0)
            {
                content.Add(descriptor);
            }
            return content;
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

        private bool IsClausesCoversTrueFalse()
        {
            if (ContainsBooleans)
            {
                return false;
            }
            var coversTrueFalse = false;
            if (_isClause.ContainsKey(CompareTokens.LT))
            {
                coversTrueFalse = _isClause[CompareTokens.LT].Any(cl => cl.CompareTo(_falseValue) > 0);
            }
            if (!coversTrueFalse && _isClause.ContainsKey(CompareTokens.GT))
            {
                coversTrueFalse = _isClause[CompareTokens.GT].Any(cl => cl.CompareTo(_trueValue) < 0);
            }
            return coversTrueFalse;
        }

        private T GetIsLTValue()
        {
            return _isClause[CompareTokens.LT].Max();
        }

        private T GetIsGTValue()
        {
            return _isClause[CompareTokens.GT].Min();
        }

        private string GetSinglesDescriptor()
        {
            var series = string.Empty;
            if (_singleValues.Any())
            {
                foreach (var val in _singleValues)
                {
                    series = series + val.ToString() + ",";
                }
                return $"Single={series.Substring(0, series.Length - 1)}";
            }
            return series;
        }

        private string GetRelOpDescriptor()
        {
            var series = string.Empty;
            if (_relationalOps.Any())
            {
                foreach (var val in _relationalOps)
                {
                    series = series + val.ToString() + ",";
                }
                return $"RelOp={series.Substring(0, series.Length - 1)}";
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

        private string GetIsLTClausesDescriptor()
        {
            return GetIsClausesDescriptor(CompareTokens.LT, "IsLT=");
        }

        private string GetIsGTClausesDescriptor()
        {
            return GetIsClausesDescriptor(CompareTokens.GT, "IsGT=");
        }

        private string GetIsClausesDescriptor(string opSymbol, string prefix)
        {
            var result = string.Empty;
            if (_isClause.TryGetValue(opSymbol, out List<T> values))
            {
                var isLT = opSymbol.Equals(CompareTokens.LT);
                var value = isLT ? values.Max() : values.Min();
                var extentToCompare = isLT ? _minExtent : _maxExtent;
                if (!(_hasExtents && value.CompareTo(extentToCompare) == 0))
                {
                    result = $"{prefix}{value.ToString()}";
                }
            }
            return result;
        }
    }

}
