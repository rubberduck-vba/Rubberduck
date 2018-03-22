using Antlr4.Runtime;
using Antlr4.Runtime.Tree;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

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

    public interface ISummaryCoverage
    {
        bool HasCoverage { get; }
        bool CoversAllValues { get; }
        bool CoversTrueFalse { get; }
        string TypeName { set; get; }
        bool TryFilterOutRedundateClauses(ISummaryCoverage summary, ref ISummaryCoverage diff);
        void Add(ISummaryCoverage newSummary);
        void AddValueRange(IUnreachableCaseInspectionValue startVal, IUnreachableCaseInspectionValue endVal);
        void AddIsClause(IUnreachableCaseInspectionValue value, string opSymbol);
        void AddSingleValue(IUnreachableCaseInspectionValue value);
        void AddRelationalOp(IUnreachableCaseInspectionValue value);
        void AddExtents(IUnreachableCaseInspectionValue minValue, IUnreachableCaseInspectionValue maxValue);
        ISummaryCoverage GetDifference(ISummaryCoverage summary);
    }

        //TODO: this interface should probably transact in uciValues
    public interface ISummaryCoverageElements<T>
    {
        bool TryGetIsLTClause(out T isLT);
        bool TryGetIsGTClause(out T isGT);
        void RemoveIsLTClause();
        void RemoveIsGTClause();
        List<Tuple<T, T>> RangeValues { get; }
        HashSet<long> DiscreteValues { get; }
        void RemoveRangeValues(List<Tuple<T, T>> toRemove);
        List<string> RelationalOps { get; }
        HashSet<T> SingleValues { get; }
    }

    public class SummaryCoverage<T> : ISummaryCoverage, ISummaryClause<T> where T : IComparable<T>
    {
        private readonly IUnreachableCaseInspectionSummaryClauseFactory _factory;
        private readonly IUnreachableCaseInspectionValueFactory _valueFactory;
        private bool _coversAllValues;
        private T _trueValue;
        private T _falseValue;

        public string TypeName {set; get;}
        public List<ParserRuleContext> IncompatibleTypeRangeContexts { set; get; } = new List<ParserRuleContext>();
        public Func<IUnreachableCaseInspectionValue, T> TConverter { private set; get; }

        private static Dictionary<string, Action<SummaryCoverage<T>, T>> IsClauseAdders = new Dictionary<string, Action<SummaryCoverage<T>, T>>()
        {
            [CompareTokens.LT] = delegate (SummaryCoverage<T> thisSum, T result) { thisSum.SetIsLT(result); },
            [CompareTokens.LTE] = delegate (SummaryCoverage<T> thisSum, T result) { thisSum.SetIsLT(result); thisSum.Add(result); },
            [CompareTokens.GT] = delegate (SummaryCoverage<T> thisSum, T result) { thisSum.SetIsGT(result); },
            [CompareTokens.GTE] = delegate (SummaryCoverage<T> thisSum, T result) { thisSum.SetIsGT(result); thisSum.Add(result); },
            [CompareTokens.EQ] = delegate (SummaryCoverage<T> thisSum, T result) { thisSum.Add(result); },
            [CompareTokens.NEQ] = delegate (SummaryCoverage<T> thisSum, T result) { thisSum.AddIsClauseNEQ(result); }
        };

        //TODO: verify the true/false values are meaningful for SummaryClause<string>
        public SummaryCoverage(IUnreachableCaseInspectionSummaryClauseFactory factory, IUnreachableCaseInspectionValueFactory valueFactory, Func<IUnreachableCaseInspectionValue, T> tConverter)
        {
            _factory = factory;
            _valueFactory = valueFactory;
            _coversAllValues = false;
            CreateSummaryClauses();
            TConverter = tConverter;
            TrueValue = tConverter(_valueFactory.Create(Tokens.True));
            FalseValue = tConverter(_valueFactory.Create(Tokens.False));
        }

        private void CreateSummaryClauses()
        {
            Ranges = new SummaryClauseRanges<T>(TConverter);
            IsLT = new SummaryClauseIsLT<T>(TConverter);
            IsGT = new SummaryClauseIsGT<T>(TConverter);
            SingleValues = new SummaryClauseSingleValues<T>(TConverter);
            RelationalOps = new SummaryClauseRelationalOps<T>(SingleValues);
        }

        private List<ISummaryClause<T>>  SummaryElements => LoadSummaryElements();

        internal SummaryClauseRanges<T> Ranges { set; get; }
        internal SummaryClauseSingleValues<T> SingleValues { set; get; }
        internal SummaryClauseIsLT<T> IsLT { set; get; }
        internal SummaryClauseIsGT<T> IsGT { set; get; }
        internal SummaryClauseRelationalOps<T> RelationalOps { set; get; }
        public bool CoversAllValues
        {
            get
            {
                if (!_coversAllValues)
                {
                    _coversAllValues = CoversAll();
                }
                return _coversAllValues;
            }
        }

        public void AddValueRange(IUnreachableCaseInspectionValue start, IUnreachableCaseInspectionValue end)
        {
            var startVal = TConverter(start);
            var endVal = TConverter(end);
            AddRange(startVal, endVal);
        }

        public void AddIsClause(IUnreachableCaseInspectionValue value, string opSymbol)
        {
            var isLTValue = TConverter(value);
            AddIsClauseResult(opSymbol, isLTValue);
        }

        public void AddSingleValue(IUnreachableCaseInspectionValue value)
        {
            var theValue = TConverter(value);
            Add(theValue);
        }

        public void AddRelationalOp(IUnreachableCaseInspectionValue value)
        {
            if (value.IsConstantValue && !CoversTrueFalse)
            {
                var theValue = TConverter(value);
                Add(theValue);
            }
            else if(!CoversTrueFalse)
            {
                RelationalOps.Add(value.ValueText);
            }
        }

        //ISummaryClause
        public bool CoversTrueFalse => SummaryElements.Any(se => se.Covers(TrueValue)) && SummaryElements.Any(se => se.Covers(FalseValue));

        //ISummaryClause
        public T TrueValue
        {
            set
            {
                _trueValue = value;
                SummaryElements.ForEach(se => se.TrueValue = _trueValue);
            }
            get
            {
                return _trueValue;
            }
        }

        //ISummaryClause
        public T FalseValue
        {
            set
            {
                _falseValue = value;
                SummaryElements.ForEach(se => se.FalseValue = _falseValue);
            }
            get
            {
                return _falseValue;
            }
        }

        //ISummaryClause
        public bool HasCoverage => SummaryElements.Any(se => se.HasCoverage);

        //ISummaryClause
        public bool Covers(T value) => SummaryElements.Any(se => se.Covers(value));

        private bool Empty => !HasCoverage;
        private bool HasExtents => !(typeof(T) == typeof(bool) || typeof(T) == typeof(string));

        private static bool ContainsBooleans => typeof(T) == typeof(bool);
        private static bool ContainsIntegerNumbers => typeof(T) == typeof(long) || typeof(T) == typeof(Int32) || typeof(T) == typeof(byte);

        private List<ISummaryClause<T>> LoadSummaryElements()
        {
            return new List<ISummaryClause<T>>()
            {
                IsLT,
                IsGT,
                Ranges,
                SingleValues,
                RelationalOps
            };
        }

        public void AddExtents(IUnreachableCaseInspectionValue min, IUnreachableCaseInspectionValue max)
        {
            throw new NotImplementedException();
        }

        public void ApplyExtents(T min, T max)
        {
            IsLT.ApplyExtents(min, max);
            IsGT.ApplyExtents(min, max);
        }

        public override bool Equals(Object o)
        {
            if (!(o is SummaryCoverage<T>))
            {
                return false;
            }
            var comp = (SummaryCoverage<T>)o;

            if (HasCoverage != comp.HasCoverage)
            {
                return false;
            }

            if (Empty && comp.Empty)
            {
                return true;
            }

            var ltIsEqual = !IsLT.HasCoverage && !comp.IsLT.HasCoverage;
            if (IsLT.HasCoverage && comp.IsLT.HasCoverage)
            {
                ltIsEqual = IsLT.Value.CompareTo(comp.IsLT.Value) == 0;
            }
            var gtIsEqual = !IsGT.HasCoverage && !comp.IsGT.HasCoverage;
            if (IsGT.HasCoverage && comp.IsGT.HasCoverage)
            {
                gtIsEqual = IsGT.Value.CompareTo(comp.IsGT.Value) == 0;
            }
            var rangesAreEqual = !Ranges.HasCoverage && !comp.Ranges.HasCoverage;
            if (Ranges.HasCoverage && comp.Ranges.HasCoverage)
            {
                rangesAreEqual = Ranges.RangeClauses.Count == comp.Ranges.RangeClauses.Count
                && Ranges.RangeClauses.All(rgs => comp.Ranges.RangeClauses.Contains(rgs));
            }
            var singleValuesAreEqual = !SingleValues.HasCoverage && !comp.SingleValues.HasCoverage;
            if (SingleValues.HasCoverage && comp.SingleValues.HasCoverage)
            {
                singleValuesAreEqual = SingleValues.Values.All(sv => comp.SingleValues.Values.Contains(sv));
            }

            return ltIsEqual && gtIsEqual && rangesAreEqual && singleValuesAreEqual;
        }

        public override int GetHashCode()
        {
            return ToString().GetHashCode();
        }

        public bool TryFilterOutRedundateClauses(ISummaryCoverage summary, ref ISummaryCoverage diff)
        {
            diff = GetDifference(summary);
            return diff.HasCoverage;           
        }

        public ISummaryCoverage GetDifference(ISummaryCoverage summary)
        {
            if (!(summary is SummaryCoverage<T> tSummary))
            {
                throw new ArgumentException($"Argument is not of type SummaryCoverage<{typeof(T).ToString()}>", "summary");
            }

            if (tSummary.CoversAllValues || Empty)
            {
                return _factory.Create(TypeName, _valueFactory);
            }

            if (tSummary.Empty)
            {
                return this;
            }

            return RemoveClausesCoveredBy(this, tSummary);
        }

        internal void SetIsLT(T newVal)
        {
            IsLT.Value = newVal;
        }

        internal void SetIsGT(T newVal)
        {
            IsGT.Value = newVal;
        }

        public void AddRange(T start, T end)
        {
            var candidate = new SummaryClauseRange<T>(start, end, TConverter);
            AddRange(candidate);
        }

        private void AddRange(ISummaryClauseRange<T> range)
        {
            if (ContainsBooleans)
            {
                SingleValues.Add(range.Start);
                SingleValues.Add(range.End);
            }
            else if (IsLT.Covers(range.Start) && IsLT.Covers(range.End)
                || IsGT.Covers(range.Start) && IsGT.Covers(range.End))
            {
                return;
            }
            Ranges.Add(range);
        }

        public void Add(T value)
        {
            if (!(IsLT.Covers(value) || IsGT.Covers(value)))
            {
                SingleValues.Add(value);
            }
        }

        public  void Add(ISummaryCoverage newSummary)
        {
            if( newSummary is SummaryCoverage<T> tSummary)
            {
                Add(tSummary);
            }
            else
            {
                throw new ArgumentException("Argument not of type SummaryCoverage<T>","newSummary");
            }
        }

        private void Add(SummaryCoverage<T> newSummary)
        {
            if (!HasExtents && newSummary.HasExtents)
            {
                ApplyExtents(newSummary.IsLT.Value, newSummary.IsGT.Value);
            }

            if (newSummary.Empty)
            {
                return;
            }

            if (newSummary.IsLT.HasCoverage)
            {
                IsLT.Add(newSummary.IsLT.Value);
            }
            if (newSummary.IsGT.HasCoverage)
            {
                IsGT.Add(newSummary.IsGT.Value);
            }

            foreach (var range in newSummary.Ranges.RangeClauses)
            {
                AddRange(range);
            }

            if (ContainsBooleans)
            {
                SingleValues.Add(newSummary.SingleValues);
            }
            else
            {
                SingleValues.Add(newSummary.SingleValues.Values.Where(sv => !(IsLT.Covers(sv) || IsGT.Covers(sv))));
            }

            if (!CoversTrueFalse)
            {
                RelationalOps.Add(newSummary.RelationalOps);
            }
        }

        private bool CoversAll()
        {
            var coversAll = false;
            if (IsLT.HasCoverage && IsGT.HasCoverage)
            {
                coversAll = IsLT.Value.CompareTo(IsGT.Value) > 0
                    || IsLT.Value.CompareTo(IsGT.Value) == 0 && SingleValues.Covers(IsLT.Value)
                    || Ranges.Covers(new SummaryClauseRange<T>(IsLT.Value, IsGT.Value, TConverter));
            }

            if (ContainsBooleans && !coversAll)
            {
                coversAll = SingleValues.Count == 2;
            }

            //For integer number types (Long, Integer, Byte), also evaluate discreet values 
            if (ContainsIntegerNumbers && !coversAll)
            {
                var allDiscreetValues = Ranges.AsIntegerNumbers;
                allDiscreetValues.AddRange(SingleValues.AsIntegerValues);

                if (IsLT.AsIntegerNumber + allDiscreetValues.Count > IsGT.AsIntegerNumber)
                {
                    var rangeStart = IsLT.AsIntegerNumber.Value + 1;
                    var rangeEnd = IsGT.AsIntegerNumber.Value - 1;
                    var valStart = _valueFactory.Create(rangeStart.ToString(), this.TypeName);
                    var valEnd = _valueFactory.Create(rangeEnd.ToString(), this.TypeName);
                    var tempRange = new SummaryClauseRange<T>(TConverter(valStart), TConverter(valEnd), TConverter);
                    coversAll = tempRange.AsIntegerNumbers.All(tv => allDiscreetValues.Contains(tv));
                    //var tempRange = new SummaryClauseRange<long>(isLTValue.Value + 1, isGTValue.Value - 1, UCIValueConverter.ConvertLong);
                    //coversAll = tempRange.AsIntegerNumbers.All(tv => allDiscreetValues.Contains(tv));
                }
            }
            return coversAll;
        }

        private void AddIsClauseResult(string compareOperator, T result)
        {
            Debug.Assert(IsClauseAdders.ContainsKey(compareOperator), "Unrecognized comparison symbol for Is Clause");
            IsClauseAdders[compareOperator](this, result);
        }

        private void AddIsClauseNEQ(T result)
        {
            if (ContainsBooleans)
            {
                if(result.ToString() == TrueValue.ToString())
                {
                    SingleValues.Add(FalseValue);
                }
                else
                {
                    SingleValues.Add(TrueValue);
                }
            }
            else
            {
                SetIsLT(result);
                SetIsGT(result);
            }
        }

        public override string ToString()
        {
            var result = string.Empty;
            result = AddToStringContent(result, IsLT.ToString());
            result = AddToStringContent(result, IsGT.ToString());
            result = AddToStringContent(result, Ranges.ToString());
            result = AddToStringContent(result, SingleValues.ToString());
            result = AddToStringContent(result, RelationalOps.ToString());
            return result;
        }

        private static string AddToStringContent(string starting, string toAdd)
        {
            if(toAdd.Length == 0)
            {
                return starting;
            }
            return starting.Length > 0 ? $"{starting}!{toAdd}" : $"{toAdd}";
        }

        private static ISummaryCoverage RemoveClausesCoveredBy(SummaryCoverage<T> removeFrom, SummaryCoverage<T> removalSpec)
        {
            var newSummary = RemoveIsClausesCoveredBy(removeFrom, removalSpec);
            newSummary = RemoveRangesCoveredBy(removeFrom, removalSpec);
            newSummary = RemoveSingleValuesCoveredBy(removeFrom, removalSpec);
            return RemoveRelationalOpsCoveredBy(removeFrom, removalSpec);
        }

        private static SummaryCoverage<T> RemoveIsClausesCoveredBy(SummaryCoverage<T> removeFrom, SummaryCoverage<T> removalSpec)
        {
            removeFrom.IsLT.ClearIfCoveredBy(removalSpec.IsLT);
            removeFrom.IsGT.ClearIfCoveredBy(removalSpec.IsGT);
            return removeFrom;
        }

        private static SummaryCoverage<T> RemoveRangesCoveredBy(SummaryCoverage<T> removeFrom, SummaryCoverage<T> removalSpec)
        {
            var toRemove = removeFrom.Ranges.RangeClauses.Where(rg => removalSpec.IsLT.Covers(rg.Start) && removalSpec.IsLT.Covers(rg.End)
                    || removalSpec.IsGT.Covers(rg.Start) && removalSpec.IsGT.Covers(rg.End)).ToList();

            removeFrom.Ranges.Remove(toRemove);

            removeFrom.Ranges.RemoveIfCoveredBy(removalSpec.Ranges);
            return removeFrom;
        }

        private static SummaryCoverage<T> RemoveSingleValuesCoveredBy(SummaryCoverage<T> removeFrom, SummaryCoverage<T> removalSpec)
        {
            List<T> toRemove = new List<T>();
            toRemove.AddRange(removeFrom.SingleValues.Values.Where(sv => removalSpec.IsLT.Covers(sv) || removalSpec.IsGT.Covers(sv)).ToList());

            foreach (var range in removalSpec.Ranges.RangeClauses)
            {
                toRemove.AddRange(removeFrom.SingleValues.Values.Where(sv => range.Covers(sv)));
            }
            removeFrom.SingleValues.Remove(toRemove);

            removeFrom.SingleValues.RemoveIfCoveredBy(removalSpec.SingleValues);

            return removeFrom;
        }

        private static ISummaryCoverage RemoveRelationalOpsCoveredBy(SummaryCoverage<T> removeFrom, SummaryCoverage<T> removalSpec)
        {
            if (removalSpec.CoversTrueFalse)
            {
                removeFrom.RelationalOps.Clear();
            }

           return removeFrom;
        }
    }

    public class SummaryCoverage2<T> : ISummaryCoverage, ISummaryCoverageElements<T> where T : IComparable<T>
    {
        private readonly IUnreachableCaseInspectionValueFactory _valueFactory;
        private readonly Func<IUnreachableCaseInspectionValue, T> _tConverter;
        private readonly T _trueValue;
        private readonly T _falseValue;

        private List<Tuple<T, T>> _ranges;
        private Dictionary<string, List<T>> _isClause;
        private HashSet<T> _singleValues;
        private HashSet<long> _discreteRangeValues;
        private List<string> _relationalOps;

        private T _minExtent;
        bool _hasExtents;
        private T _maxExtent;

        private static bool ContainsBooleans => typeof(T) == typeof(bool);
        private static bool ContainsIntegerNumbers => typeof(T) == typeof(long) || typeof(T) == typeof(Int32) || typeof(T) == typeof(byte);

        public SummaryCoverage2(IUnreachableCaseInspectionValueFactory valueFactory, Func<IUnreachableCaseInspectionValue, T> tConverter)
        {
            _valueFactory = valueFactory;
            _tConverter = tConverter;

            _ranges = new List<Tuple<T, T>>();
            _singleValues = new HashSet<T>();
            _isClause = new Dictionary<string, List<T>>();
            _relationalOps = new List<string>();
            _discreteRangeValues = new HashSet<long>();
            _hasExtents = false;
            _trueValue = _tConverter(_valueFactory.Create("True", TypeName));
            _falseValue = _tConverter(_valueFactory.Create("False", TypeName));
        }

        //ISummaryCoverage
        public bool CoversAllValues
        {
            get
            {
                var coversAll = false;
                if (_isClause.ContainsKey(CompareTokens.LT) && _isClause.ContainsKey(CompareTokens.GT))
                {
                    coversAll = GetIsLTValue().CompareTo(GetIsGTValue()) > 0
                        || GetIsLTValue().CompareTo(GetIsGTValue()) == 0 && SingleValues.Contains(GetIsLTValue())
                        || RangeValues.Any(rv => rv.Item1.CompareTo(GetIsLTValue()) <= 0 && rv.Item2.CompareTo(GetIsGTValue()) >= 0);
                }

                if (ContainsBooleans && !coversAll)
                {
                    //TODO: Add test and code for IsLT = 1 and IsGT = -1
                    coversAll = SingleValues.Count == 2;
                }
                //For integer number types (Long, Integer, Byte), also evaluate discreet values 
                if (ContainsIntegerNumbers && !coversAll)
                {
                    if (_isClause.ContainsKey(CompareTokens.LT) && _isClause.ContainsKey(CompareTokens.GT))
                    {
                        var allDiscreetValues = RangesAsIntegerNumbers();
                        allDiscreetValues.AddRange(SingleValues.Select(sv => long.Parse(sv.ToString())));
                        if (long.Parse(GetIsLTValue().ToString()) + allDiscreetValues.Count > long.Parse(GetIsGTValue().ToString()))
                        {
                            var rangeStart = long.Parse(GetIsLTValue().ToString()) + 1;
                            var rangeEnd = long.Parse(GetIsGTValue().ToString()) - 1;
                            var remainingValues = new List<long>();
                            for (var idx = rangeStart; idx <= rangeEnd; idx++)
                            {
                                remainingValues.Add(idx);
                            }
                            coversAll = remainingValues.All(rv => allDiscreetValues.Contains(rv));
                        }
                    }
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
                // return false;
                return _ranges.Any()
                    || _singleValues.Any()
                    || _isClause.Any()
                    || _relationalOps.Any();
            }
        }

        //ISummaryCoverage
        public void Add(ISummaryCoverage newSummary)
        {
            var itf = (ISummaryCoverageElements<T>)newSummary;
            if (itf.TryGetIsLTClause(out T isLT))
            {
                AddIsClauseImpl(isLT, CompareTokens.LT);
            }
            if (itf.TryGetIsGTClause(out T isGT))
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
                AddRelationalOp(_valueFactory.Create(op,TypeName));
            }
            var singleVals = itf.SingleValues;
            foreach (var val in singleVals)
            {
                AddSingleValueImpl(val);
                //AddSingleValue(_valueFactory.Create(val.ToString(), TypeName));
            }
        }

        //ISummaryCoverage
        public void AddExtents(IUnreachableCaseInspectionValue min, IUnreachableCaseInspectionValue max)
        {
            _hasExtents = true;
            _minExtent = _tConverter(min);
            _maxExtent = _tConverter(max);
            AddIsClause(min, CompareTokens.LT);
            AddIsClause(max, CompareTokens.GT);
        }

        //ISummaryCoverage
        public void AddIsClause(IUnreachableCaseInspectionValue value, string opSymbol)
        {
            if (ContainsBooleans)
            {
                //TODO: Introduce truth table?
                return;
            }

            if(opSymbol.Equals(CompareTokens.LT) || opSymbol.Equals(CompareTokens.GT))
            {
                if (!_isClause.Keys.Contains(opSymbol))
                {
                    _isClause.Add(opSymbol, new List<T>());
                }
                _isClause[opSymbol].Add(_tConverter(value));
            }
            else if (opSymbol.Equals(CompareTokens.LTE) || opSymbol.Equals(CompareTokens.GTE))
            {
                var ltOrGtSymbol = opSymbol.Substring(0, opSymbol.Length - 1);
                if (!_isClause.Keys.Contains(ltOrGtSymbol))
                {
                    _isClause.Add(ltOrGtSymbol, new List<T>());
                }
                _isClause[ltOrGtSymbol].Add(_tConverter(value));
                AddSingleValue(value);
            }
            else if (opSymbol.Equals(CompareTokens.EQ))
            {
                AddSingleValue(value);
            }
            else if (opSymbol.Equals(CompareTokens.NEQ))
            {
                _isClause[CompareTokens.LT].Add(_tConverter(value));
                _isClause[CompareTokens.GT].Add(_tConverter(value));
            }
        }

        //ISummaryCoverage
        public void AddRelationalOp(IUnreachableCaseInspectionValue value)
        {
            if (value.IsConstantValue)
            {
                AddSingleValueImpl(_tConverter(value));
            }
            else
            {
                AddRelationalOpImpl(value.ValueText);
            }
        }

        //ISummaryCoverage
        public void AddSingleValue(IUnreachableCaseInspectionValue value)
        {
            AddSingleValueImpl(_tConverter(value));
        }

        //ISummaryCoverage
        public void AddValueRange(IUnreachableCaseInspectionValue inputStartVal, IUnreachableCaseInspectionValue inputEndVal)
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
        public bool TryFilterOutRedundateClauses(ISummaryCoverage existingCoverage, ref ISummaryCoverage augmentingCoverage)
        {
            if (!(existingCoverage is SummaryCoverage2<T>))
            {
                throw new ArgumentException($"Argument is not of type SummaryCoverage<{typeof(T).ToString()}>", "summary");
            }

            if (!(augmentingCoverage is SummaryCoverage2<T>))
            {
                throw new ArgumentException($"Argument is not of type SummaryCoverage<{typeof(T).ToString()}>", "summary");
            }

            if (!HasCoverage || existingCoverage.CoversAllValues)
            {
                return false;
            }

            if (!existingCoverage.HasCoverage)
            {
                augmentingCoverage = this;
                return true;
            }

            augmentingCoverage = RemoveClausesCoveredBy(this, (ISummaryCoverageElements<T>)existingCoverage);
            return augmentingCoverage.HasCoverage;
        }

        //ISummaryCoverage
        public ISummaryCoverage GetDifference(ISummaryCoverage summary)
        {
            throw new NotImplementedException();
        }

        private void AddSingleValueImpl(T value)
        {
            _singleValues.Add(value);
        }

        private bool CoversRange(T start, T end)
        {
            if (!_ranges.Any())
            {
                return false;
            }
            var existingCoversProposed = _ranges.Where(t => t.Item1.CompareTo(start) <= 0 && t.Item2.CompareTo(end) >= 0);
            return existingCoversProposed.Any();
        }

        //TODO: Get this impl to handle the AddIsClause call at the bottom 
        //of the function
        private void AddIsClauseImpl(T val, string opSymbol)
        {
            if (ContainsBooleans)
            {
                //TODO: introduce Truth table of observed behavior and write to SingleValues
                return;
            }
            var value = _valueFactory.Create(val.ToString(), TypeName);
            AddIsClause(value, opSymbol);
        }

        private void AddRelationalOpImpl(string value)
        {
            if (!CoversTrueFalse)
            {
                _relationalOps.Add(value);
                return;
            }
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

            if (!_ranges.Any())
            {
                _ranges.Add(new Tuple<T, T>(start, end));
                LoadDiscretes();
                return;
            }

            if (CoversRange(start, end))
            {
                return;
            }

            bool rangeIsAdded = TryMergeWithExistingRanges(start, end);


            if (!rangeIsAdded)
            {
                _ranges.Add(new Tuple<T, T>(start, end));
                LoadDiscretes();
            }
        }

        private void LoadDiscretes()
        {
            if (ContainsIntegerNumbers)
            {
                foreach (var range in _ranges)
                {
                    var rangeStart = long.Parse(range.Item1.ToString());
                    var rangeEnd = long.Parse(range.Item2.ToString());
                    for (var val = rangeStart; val <= rangeEnd; val++)
                    {
                        _discreteRangeValues.Add(val);
                    }
                }
            }
        }

        private bool TryMergeWithExistingRanges(T start, T end)
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
                    LoadDiscretes();
                    rangeIsAdded = true;
                }
                else
                {
                    var original = startIsWithin.First();
                    _ranges.Remove(startIsWithin.First());
                    _ranges.Add(new Tuple<T, T>(original.Item1, end));
                    LoadDiscretes();
                    rangeIsAdded = true;
                }
            }

            return rangeIsAdded;
        }

        //public ISummaryCoverage GetDifference(ISummaryCoverage summary)
        //{
        //    throw new NotImplementedException();
        //}

        //ISummaryCoverageElements<T>

        public HashSet<long> DiscreteValues => _discreteRangeValues;

        public void RemoveIsLTClause()
        {
            if (_isClause.Keys.Contains(CompareTokens.LT))
            {
                _isClause.Remove(CompareTokens.LT);
            }
        }

        //ISummaryCoverageElements<T>
        public bool TryGetIsLTClause(out T isLT)
        {
            isLT = default;
            var clauses = new Dictionary<string, T>();

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

        //ISummaryCoverageElements<T>
        public void RemoveIsGTClause()
        {
            if (_isClause.Keys.Contains(CompareTokens.GT))
            {
                _isClause.Remove(CompareTokens.GT);
            }
        }

        //ISummaryCoverageElements<T>
        public bool TryGetIsGTClause(out T isGT)
        {
            isGT = default;
            var clauses = new Dictionary<string, T>();

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

        //ISummaryCoverageElements<T>
        public List<Tuple<T, T>> RangeValues => _ranges;

        //ISummaryCoverageElements<T>
        public void RemoveRangeValues(List<Tuple<T, T>> toRemove)
        {
            foreach(var tp in toRemove)
            {
                _ranges.Remove(tp);
            }
        }

        //ISummaryCoverageElements<T>
        public List<string> RelationalOps => _relationalOps;

        //ISummaryCoverageElements<T>
        public HashSet<T> SingleValues => _singleValues;

        public override string ToString()
        {
            var descriptors = new List<string>();
            descriptors = AddDesciptor(GetIsLTClausesDescriptor(), descriptors);
            descriptors = AddDesciptor(GetIsGTClausesDescriptor(), descriptors);
            descriptors = AddDesciptor(GetRangesDescriptor(), descriptors);
            descriptors = AddDesciptor(GetSinglesDescriptor(), descriptors);
            descriptors = AddDesciptor(GetRelOpDescriptor(), descriptors);
            var descriptor = string.Empty;
            foreach( var desc in descriptors)
            {
                descriptor = descriptor + desc + "!";
            }
            if(descriptor.Length > 0)
            {
                return descriptor.Substring(0, descriptor.Length - 1);
            }
            return string.Empty;
        }

        private static List<string> AddDesciptor(string descriptor, List<string> content)
        {
            if (descriptor.Length > 0)
            {
                content.Add(descriptor);
            }
            return content;
        }

        private List<long> RangesAsIntegerNumbers()
        {
            var results = new List<long>();
            if (ContainsIntegerNumbers)
            {
                foreach (var range in RangeValues)
                {
                    var start = long.Parse(range.Item1.ToString());
                    var end = long.Parse(range.Item2.ToString());
                    for (var idx = start; idx <= end; idx++)
                    {
                        results.Add(idx);
                    }
                }
            }
            return results;
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
            if(_isClause.TryGetValue(opSymbol, out List<T> values))
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



        private static ISummaryCoverage RemoveClausesCoveredBy(ISummaryCoverageElements<T> removeFrom, ISummaryCoverageElements<T> removalSpec)
        {
            var newSummary = RemoveIsClausesCoveredBy(removeFrom, removalSpec);
            newSummary = RemoveRangesCoveredBy(removeFrom, removalSpec);
            newSummary = RemoveSingleValuesCoveredBy(removeFrom, removalSpec);
            newSummary = RemoveRelationalOpsCoveredBy(removeFrom, removalSpec);
            return (ISummaryCoverage)newSummary;
        }

        private static ISummaryCoverageElements<T> RemoveIsClausesCoveredBy(ISummaryCoverageElements<T> removeFrom, ISummaryCoverageElements<T> removalSpec)
        {
            if (removeFrom.TryGetIsLTClause(out T isLT))
            {
                if (removalSpec.TryGetIsLTClause(out T removalSpecLT))
                {
                    if (removalSpecLT.CompareTo(isLT) >= 0)
                    {
                        removeFrom.RemoveIsLTClause();
                    }
                }
            }
            if (removeFrom.TryGetIsGTClause(out T isGT))
            {
                if (removalSpec.TryGetIsGTClause(out T removalSpecGT))
                {
                    if (removalSpecGT.CompareTo(isGT) <= 0)
                    {
                        removeFrom.RemoveIsGTClause();
                    }
                }
            }
            return removeFrom;
        }

        private static ISummaryCoverageElements<T> RemoveRangesCoveredBy(ISummaryCoverageElements<T> removeFrom, ISummaryCoverageElements<T> removalSpec)
        {
            if (!removeFrom.RangeValues.Any())
            {
                return removeFrom;
            }

            var rangesToRemove = new List<Tuple<T, T>>();
            if (removalSpec.TryGetIsLTClause(out T removalSpecLT))
            {
                foreach (var tup in removeFrom.RangeValues)
                {
                    if (removalSpecLT.CompareTo(tup.Item1) > 0 && removalSpecLT.CompareTo(tup.Item2) > 0)
                    {
                        rangesToRemove.Add(tup);
                    }
                }
            }

            if (removalSpec.TryGetIsGTClause(out T removalSpecGT))
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
                foreach(var rem in removalSpec.RangeValues)
                {
                    if(rem.Item1.CompareTo(tup.Item1) <= 0 && rem.Item2.CompareTo(tup.Item2) >= 0)
                    {
                        rangesToRemove.Add(tup);
                    }
                }
            }
            removeFrom.RemoveRangeValues(rangesToRemove);

            rangesToRemove.Clear();
            //var contained = new List<bool>();
            var canBeFiltered = true;
            if (ContainsIntegerNumbers && removalSpec.DiscreteValues.Any())
            {
                foreach(var rem in removeFrom.RangeValues)
                {
                    var rangeStart = long.Parse(rem.Item1.ToString());
                    var rangeEnd = long.Parse(rem.Item2.ToString());
                    for (var val = rangeStart; val <= rangeEnd && canBeFiltered; val++)
                    {
                        if (!removalSpec.DiscreteValues.Contains(val))
                        {
                            canBeFiltered = false; ;
                        }
                    }
                    if (canBeFiltered)
                    {
                        rangesToRemove.Add(rem);
                    }
                }
                removeFrom.RemoveRangeValues(rangesToRemove);
            }

            return removeFrom;
        }

        private static ISummaryCoverageElements<T> RemoveSingleValuesCoveredBy(ISummaryCoverageElements<T> removeFrom, ISummaryCoverageElements<T> removalSpec)
        {
            List<T> toRemove = new List<T>();
            if (removalSpec.TryGetIsLTClause(out T removalSpecLT))
            {
                foreach (var sv in removeFrom.SingleValues)
                {
                    if (removalSpecLT.CompareTo(sv) > 0)
                    {
                        toRemove.Add(sv);
                    }
                }
            }

            if (removalSpec.TryGetIsGTClause(out T removalSpecGT))
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

            foreach(var rem in toRemove)
            {
                removeFrom.SingleValues.Remove(rem);
            }
            return removeFrom;
        }

        private static ISummaryCoverageElements<T> RemoveRelationalOpsCoveredBy(ISummaryCoverageElements<T> removeFrom, ISummaryCoverageElements<T> removalSpec)
        {
            List<string> toRemove = new List<string>();
            if (((ISummaryCoverage)removalSpec).CoversTrueFalse)
            {
                removeFrom.RelationalOps.Clear();
            }
            foreach(var rem in removalSpec.RelationalOps)
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
    }

}
