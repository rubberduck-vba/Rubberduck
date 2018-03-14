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
        bool CoversAllValues { get; }
        bool CoversTrueFalse { get; }
        void Add(ISummaryCoverage newSummary);
        string TypeName { set; get; }
        bool HasCoverage { get; }
        bool HasExtents { get; }
        bool CanBeInspected { get; }
        bool HasClausesNotCoveredBy(ISummaryCoverage summaryCoverage, out ISummaryCoverage diff);
        void AddValueRange(IUnreachableCaseInspectionValue startVal, IUnreachableCaseInspectionValue endVal);
        void AddIsClause(IUnreachableCaseInspectionValue value, string opSymbol);
        void AddSingleValue(IUnreachableCaseInspectionValue value);
        void AddRelationalOp(IUnreachableCaseInspectionValue value);
    }

    public interface IUnreachablCaseValueConverter<T>
    {
        Func<IUnreachableCaseInspectionValue, T> TConverter { set; }
    }

    public class SummaryCoverage<T> : ISummaryCoverage, ISummaryClause<T>, IUnreachablCaseValueConverter<T> where T : System.IComparable<T>
    {
        private readonly ISummaryCoverageFactory _factory;
        private bool _coversAllValues;
        private T _trueValue;
        private T _falseValue;

        public Func<IUnreachableCaseInspectionValue, T> TConverter { set; get; }

        public string TypeName {set; get;}
        public List<ParserRuleContext> IncompatibleTypeRangeContexts { set; get; } = new List<ParserRuleContext>();

        private static Dictionary<string, Action<SummaryCoverage<T>, T>> IsClauseAdders = new Dictionary<string, Action<SummaryCoverage<T>, T>>()
        {
            [CompareTokens.LT] = delegate (SummaryCoverage<T> thisSum, T result) { thisSum.SetIsLT(result); },
            [CompareTokens.LTE] = delegate (SummaryCoverage<T> thisSum, T result) { thisSum.SetIsLT(result); thisSum.Add(result); },
            [CompareTokens.GT] = delegate (SummaryCoverage<T> thisSum, T result) { thisSum.SetIsGT(result); },
            [CompareTokens.GTE] = delegate (SummaryCoverage<T> thisSum, T result) { thisSum.SetIsGT(result); thisSum.Add(result); },
            [CompareTokens.EQ] = delegate (SummaryCoverage<T> thisSum, T result) { thisSum.Add(result); },
            [CompareTokens.NEQ] = delegate (SummaryCoverage<T> thisSum, T result) { thisSum.AddIsClauseNEQ(result); }
        };

        public SummaryCoverage(ISummaryCoverageFactory factory, T min, T max, T trueVal, T falseVal)
        {
            _factory = factory;
            _coversAllValues = false;
            if (HasExtents)
            {
                ApplyExtents(min, max);
            }
            RelationalOps = new SummaryClauseRelationalOps<T>(SingleValues);
            TrueValue = trueVal;
            FalseValue = falseVal;
        }

        public SummaryCoverage(ISummaryCoverageFactory factory, T trueVal, T falseVal)
        {
            _factory = factory;
            _coversAllValues = false;
            RelationalOps = new SummaryClauseRelationalOps<T>(SingleValues);
            TrueValue = trueVal;
            FalseValue = falseVal;
        }

        private List<ISummaryClause<T>>  SummaryElements => LoadSummaryElements();

        public Dictionary<ParserRuleContext, SummaryCoverage<T>> RangeClauseSummaries { set; get; } = new Dictionary<ParserRuleContext, SummaryCoverage<T>>();
        public SummaryClauseRanges<T> Ranges { set; get; } = new SummaryClauseRanges<T>();
        public SummaryClauseSingleValues<T> SingleValues { set; get; } = new SummaryClauseSingleValues<T>();
        public SummaryClauseIsLT<T> IsLT { set; get; } = new SummaryClauseIsLT<T>();
        public SummaryClauseIsGT<T> IsGT { set; get; } = new SummaryClauseIsGT<T>();
        public SummaryClauseRelationalOps<T> RelationalOps { set; get; }
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

        private static Dictionary<ParserRuleContext, bool> LoadResult(ParserRuleContext ctxt, bool result, Dictionary<ParserRuleContext, bool> container)
        {
            if (container.ContainsKey(ctxt))
            {
                if(result && container[ctxt] != result)
                {
                    container[ctxt] = result;
                }
            }
            else
            {
                container.Add(ctxt, result);
            }
            return container;
        }

        public bool CanBeInspected => HasCoverage || HasExtents;


        //public bool CanBeInspected(VBAParser.CaseClauseContext caseClause)
        //{
        //    //var ranges = caseClause.rangeClause();
        //    //var summariesOfInterest = RangeClauseSummaries.Where(rgs => ranges.Contains(rgs.Key)).Select(sum => sum);
        //    //var coverage = summariesOfInterest.Any(sum => sum.Value.HasCoverage || sum.Value.HasExtents);
        //    var coverage = HasCoverage || HasExtents;
        //    //return coverage && !IsIncompatibleType(caseClause);
        //    return coverage;
        //}

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

        public bool Empty => !HasCoverage;
        public bool HasExtents => !(typeof(T) == typeof(bool) || typeof(T) == typeof(string));

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

        public void ApplyExtents(T min, T max)
        {
            //_extents.MinMax(min, max);
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

        public bool HasClausesNotCoveredBy(ISummaryCoverage summaryCoverage, out ISummaryCoverage diff)
        {
            if(!(summaryCoverage is SummaryCoverage<T> tSummary))
            {
                throw new ArgumentException("Argument is not of type SummaryCoverage<T>", "summaryCoverage");
            }

            diff = CreateSummaryCoverageDifference(tSummary);
            return diff.HasCoverage;
        }

        private ISummaryCoverage CreateSummaryCoverageDifference(SummaryCoverage<T> toRemove)
        {
            if (toRemove.CoversAllValues || Empty)
            {
                return _factory.Create(TypeName);
            }

            if (toRemove.Empty)
            {
                return this;
            }

            return RemoveClausesCoveredBy(this, toRemove);
        }

        public void SetIsLT(T newVal)
        {
            IsLT.Value = newVal;
        }

        public void SetIsGT(T newVal)
        {
            IsGT.Value = newVal;
        }

        public void AddRange(T start, T end)
        {
            var candidate = new SummaryClauseRange<T>(start, end);
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

        //public ISummaryCoverage CoverageForCaseClause(VBAParser.CaseClauseContext caseClause)
        //{
        //    var caseClauseCoverage = _factory.Create(this.TypeName);
        //    foreach (var range in caseClause.rangeClause())
        //    {
        //        caseClauseCoverage.Add(CoverageForRangeClause(range));
        //    }
        //    return caseClauseCoverage;
        //}

        public ISummaryCoverage CoverageForCaseClauseX(VBAParser.CaseClauseContext caseClause)
        {
            //var caseClause = obj as VBAParser.CaseClauseContext;
            var caseClauseCoverage = _factory.Create(this.TypeName);
            foreach (var range in caseClause.rangeClause())
            {
                caseClauseCoverage.Add(CoverageForRangeClauseX(range));
            }
            return caseClauseCoverage;
        }

        //public ISummaryCoverage CoverageForRangeClause(VBAParser.RangeClauseContext context)
        //{
        //    if (RangeClauseSummaries.ContainsKey(context))
        //    {
        //        return RangeClauseSummaries[context];
        //    }
        //    return _factory.Create(this.TypeName);
        //}

        public ISummaryCoverage CoverageForRangeClauseX(VBAParser.RangeClauseContext rangeClause)
        {
            //var context = rangeClause as VBAParser.RangeClauseContext;
            if (RangeClauseSummaries.ContainsKey(rangeClause))
            {
                return RangeClauseSummaries[rangeClause];
            }
            return _factory.Create(this.TypeName);
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
                //ApplyExtents(newSummary.Extents.Min, newSummary.Extents.Max);
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
                    || Ranges.Covers(new SummaryClauseRange<T>(IsLT.Value, IsGT.Value));
            }

            if (ContainsBooleans && !coversAll)
            {
                coversAll = SingleValues.Count == 2;
            }

            if (ContainsIntegerNumbers && !coversAll)
            {
                var allDiscreetValues = Ranges.AsIntegerNumbers;
                allDiscreetValues.AddRange(SingleValues.AsIntegerValues);

                long? isLTValue = IsLT.AsIntegerNumber;
                long? isGTValue = IsGT.AsIntegerNumber;

                if (isLTValue + allDiscreetValues.Count > isGTValue)
                {
                    var tempRange = new SummaryClauseRange<long>(isLTValue.Value + 1, isGTValue.Value - 1);
                    coversAll = tempRange.AsIntegerNumbers.All(tv => allDiscreetValues.Contains(tv));
                }
            }
            return coversAll;
        }

        //////Used to modify logic operators to convert LHS and RHS for expressions like '5 > x' (= 'x < 5')
        //public static Dictionary<string, string> AlgebraicLogicalInversions = new Dictionary<string, string>()
        //{
        //    [CompareTokens.EQ] = CompareTokens.EQ,
        //    [CompareTokens.NEQ] = CompareTokens.NEQ,
        //    [CompareTokens.LT] = CompareTokens.GT,
        //    [CompareTokens.LTE] = CompareTokens.GTE,
        //    [CompareTokens.GT] = CompareTokens.LT,
        //    [CompareTokens.GTE] = CompareTokens.LTE
        //};

        //public static Dictionary<string, Func<ParseTreeValue, ParseTreeValue, ParseTreeValue>>
        //    BinaryLogicalOps = new Dictionary<string, Func<ParseTreeValue, ParseTreeValue, ParseTreeValue>>()
        //    {
        //        [CompareTokens.GT] = (LHS, RHS) => LHS > RHS ? ParseTreeValue.True : ParseTreeValue.False,
        //        [CompareTokens.GTE] = (LHS, RHS) => LHS >= RHS ? ParseTreeValue.True : ParseTreeValue.False,
        //        [CompareTokens.LT] = (LHS, RHS) => LHS < RHS ? ParseTreeValue.True : ParseTreeValue.False,
        //        [CompareTokens.LTE] = (LHS, RHS) => LHS <= RHS ? ParseTreeValue.True : ParseTreeValue.False,
        //        [CompareTokens.EQ] = (LHS, RHS) => LHS == RHS ? ParseTreeValue.True : ParseTreeValue.False,
        //        [CompareTokens.NEQ] = (LHS, RHS) => LHS != RHS ? ParseTreeValue.True : ParseTreeValue.False,
        //        [Tokens.And] = (LHS, RHS) => LHS.AsBoolean().Value && RHS.AsBoolean().Value ? ParseTreeValue.True : ParseTreeValue.False,
        //        [Tokens.Or] = (LHS, RHS) => LHS.AsBoolean().Value || RHS.AsBoolean().Value ? ParseTreeValue.True : ParseTreeValue.False,
        //        [Tokens.XOr] = (LHS, RHS) => LHS.AsBoolean().Value ^ RHS.AsBoolean().Value ? ParseTreeValue.True : ParseTreeValue.False,
        //        [Tokens.Not] = (LHS, RHS) => LHS.AsBoolean().Value || RHS.AsBoolean().Value ? ParseTreeValue.True : ParseTreeValue.False
        //        //["Eqv"] = (LHS, RHS) => LHS.AsBoolean().Value ^ RHS.AsBoolean().Value ? UnreachableCaseInspectionValue.True : UnreachableCaseInspectionValue.False
        //        //["Imp"] = (LHS, RHS) => LHS.AsBoolean().Value ^ RHS.AsBoolean().Value ? UnreachableCaseInspectionValue.True : UnreachableCaseInspectionValue.False,
        //    };

        //private bool RangeStartOrEndHasMismatch(ParserRuleContext prCtxt)
        //{
        //    bool mismatchFound = false;
        //    foreach (var ctxt in prCtxt.GetChildren<ParserRuleContext>())
        //    {
        //        if (ParseTreeValueResults.VariableContexts.Keys.Contains(ctxt))
        //        {
        //            var value = ParseTreeValueResults.VariableContexts[ctxt];
        //            if (!mismatchFound)
        //            {
        //                mismatchFound = !value.HasValueAs(TypeName);
        //            }
        //        }
        //    }
        //    return mismatchFound;
        //}

        //public void LoadRangeClauseCoverageX(ParserRuleContext selectStmt, Dictionary<ParserRuleContext,T> valueResolvedContexts)
        //{
        //    var rangeClauses = ((VBAParser.SelectCaseStmtContext)selectStmt).caseClause().SelectMany(cc => cc.rangeClause());
        //    foreach (ParserRuleContext rangeClause in rangeClauses)
        //    {
        //        var rangeSummaryCoverage = (SummaryCoverage<T>)_factory.Create(this.TypeName);
        //        if (rangeClause.HasChildToken(Tokens.To))
        //        {
        //            var startContext = rangeClause.GetChild<VBAParser.SelectStartValueContext>();
        //            var endContext = rangeClause.GetChild<VBAParser.SelectEndValueContext>();

        //            var hasStart = valueResolvedContexts.TryGetValue(startContext, out T startVal);
        //            var hasEnd = valueResolvedContexts.TryGetValue(endContext, out T endVal);

        //            if (hasStart && hasEnd)
        //            {
        //                rangeSummaryCoverage.AddRange(startVal, endVal);
        //            }
        //            else
        //            {
        //                var ctxt = !hasStart ? (ParserRuleContext)startContext: endContext;
        //                if (RangeStartOrEndHasMismatch(ctxt))
        //                {
        //                    IncompatibleTypeRangeContexts.Add(rangeClause);
        //                }
        //            }
        //        }
        //        else //single value
        //        {
        //            var ctxts = rangeClause.children.Where(ch => ch is ParserRuleContext
        //                            && valueResolvedContexts.Keys.Contains((ParserRuleContext)ch));

        //            //Is Statements
        //            if (ctxts.Any() && ctxts.Count() == 1 && rangeClause.HasChildToken(Tokens.Is))
        //            {
        //                var compOpContext = rangeClause.GetChild<VBAParser.ComparisonOperatorContext>();
        //                rangeSummaryCoverage.AddIsClauseResult(compOpContext.GetText(), valueResolvedContexts[(ParserRuleContext)ctxts.First()]);
        //            }
        //            //RelationalOp statements like x < 100, 100 < x
        //            else if (rangeClause.TryGetChildContext(out VBAParser.RelationalOpContext relOpCtxt))
        //            {
        //                if (valueResolvedContexts.Keys.Contains(relOpCtxt))
        //                {
        //                    rangeSummaryCoverage.RelationalOps.Add(valueResolvedContexts[relOpCtxt]);
        //                }
        //                else
        //                {
        //                    rangeSummaryCoverage.RelationalOps.Add(relOpCtxt.GetText());
        //                }
        //            }
        //            else if (ctxts.Any() && ctxts.Count() == 1)
        //            {
        //                rangeSummaryCoverage.Add(valueResolvedContexts[(ParserRuleContext)ctxts.First()]);
        //            }
        //        }
        //        RangeClauseSummaries.Add(rangeClause, rangeSummaryCoverage);
        //    }
        //}

        public void AddIsClauseResult(string compareOperator, T result)
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
}
