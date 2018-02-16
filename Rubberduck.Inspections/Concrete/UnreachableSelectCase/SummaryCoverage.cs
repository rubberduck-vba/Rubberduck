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

namespace Rubberduck.Inspections.Concrete
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
        ISummaryCoverage CoverageFor(VBAParser.RangeClauseContext context);
        ISummaryCoverage CoverageFor(VBAParser.CaseClauseContext context);
        bool HasCoverage { get; }
        bool HasExtents { get; }
        bool CanBeInspected(IEnumerable<ParserRuleContext> ranges);
        bool HasConditionsNotCoveredBy(ISummaryCoverage summaryCoverage, out ISummaryCoverage diff);
        bool IsIncompatibleType(IEnumerable<ParserRuleContext> ranges);
        IParseTreeValueResults ParseTreeValueResults { set; get; }
    }

    public class SummaryCoverage<T> : ISummaryCoverage, ISummaryClause<T> where T : System.IComparable<T>
    {
        private ContextExtents<T> _extents;
        private bool _coversAllValues;
        private T _trueValue;
        private T _falseValue;

        public string TypeName {set; get;}
        public Dictionary<ParserRuleContext, T> TypedValueResults { set; get; } = new Dictionary<ParserRuleContext, T>();
        public IParseTreeValueResults ParseTreeValueResults { set; get; } = null;

        private Dictionary<ParserRuleContext, SummaryCoverage<T>> _rangeClauseSummaries;
        private List<ParserRuleContext> IncompatibleTypeRangeContexts { set; get; } = new List<ParserRuleContext>();

        private static Dictionary<string, Action<SummaryCoverage<T>, T>> IsClauseAdders = new Dictionary<string, Action<SummaryCoverage<T>, T>>()
        {
            [CompareTokens.LT] = delegate (SummaryCoverage<T> thisSum, T result) { thisSum.SetIsLT(result); },
            [CompareTokens.LTE] = delegate (SummaryCoverage<T> thisSum, T result) { thisSum.SetIsLT(result); thisSum.Add(result); },
            [CompareTokens.GT] = delegate (SummaryCoverage<T> thisSum, T result) { thisSum.SetIsGT(result); },
            [CompareTokens.GTE] = delegate (SummaryCoverage<T> thisSum, T result) { thisSum.SetIsGT(result); thisSum.Add(result); },
            [CompareTokens.EQ] = delegate (SummaryCoverage<T> thisSum, T result) { thisSum.Add(result); },
            [CompareTokens.NEQ] = delegate (SummaryCoverage<T> thisSum, T result) { thisSum.AddIsClauseNEQ(result); }
        };

        public SummaryCoverage()
        {
            _coversAllValues = false;
            _rangeClauseSummaries = new Dictionary<ParserRuleContext, SummaryCoverage<T>>();
            ApplyExtents(new ContextExtents<T>());
            RelationalOps = new SummaryClauseRelationalOps<T>(SingleValues);
            _summaryElements = LoadSummaryElements();
        }

        public SummaryCoverage(ContextExtents<T> extents, T trueVal, T falseVal)
        {
            _coversAllValues = false;
            _rangeClauseSummaries = new Dictionary<ParserRuleContext, SummaryCoverage<T>>();
            ApplyExtents(extents);
            RelationalOps = new SummaryClauseRelationalOps<T>(SingleValues);
            _summaryElements = LoadSummaryElements();
            TrueValue = trueVal;
            FalseValue = falseVal;
        }

        private ContextExtents<T> Extents => _extents;
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
        
        public bool IsIncompatibleType(IEnumerable<ParserRuleContext> ranges)
        {
            var results = new Dictionary<ParserRuleContext, bool>();
            var summariesOfInterest = _rangeClauseSummaries.Where(rgs => ranges.Contains(rgs.Key)).Select(sum => sum);
            foreach( var summary in summariesOfInterest)
            {
                var rangeIsIncompatible = IncompatibleTypeRangeContexts.Contains(summary.Key);
                results = LoadResult(summary.Key, rangeIsIncompatible, results);

                var valueIsIncompatible = false;
                var valueResolvedContexts = ParseTreeValueResults.ValueResolvedContexts.Where(vrc => vrc.Key.IsDescendentOf(summary.Key)).Select(vrc => vrc.Value);
                foreach(var value in valueResolvedContexts)
                {
                    valueIsIncompatible = !value.HasValueAs(value.UseageTypeName);
                    results = LoadResult(summary.Key, valueIsIncompatible, results);
                }
                var variableIsIncompatible = false;
                var variableContexts = ParseTreeValueResults.VariableContexts.Where(vrc => vrc.Key.IsDescendentOf(summary.Key)).Select(vrc => vrc.Value);
                foreach (var value in variableContexts)
                {
                    variableIsIncompatible = value.DerivedTypeName != value.UseageTypeName;
                    results = LoadResult(summary.Key, variableIsIncompatible, results);
                }
            }

            return results.Values.All(v => v == true); 
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

        private bool RangeClauseIsIncompatibleType(ParserRuleContext rangeClause)
        {
            if (rangeClause.HasChildToken(Tokens.To))
            {
                return IncompatibleTypeRangeContexts.Contains(rangeClause);
            }
            var rangeConcrete = _rangeClauseSummaries[rangeClause];
            return false;
        }

        public bool CanBeInspected(IEnumerable<ParserRuleContext> ranges)
        {
            var summariesOfInterest = _rangeClauseSummaries.Where(rgs => ranges.Contains(rgs.Key)).Select(sum => sum);
            var coverage = summariesOfInterest.Any(sum => sum.Value.HasCoverage || sum.Value.HasExtents);
            return coverage && !IsIncompatibleType(ranges);
        }

        public bool CoversTrueFalse => _summaryElements.Any(se => se.Covers(TrueValue)) && _summaryElements.Any(se => se.Covers(FalseValue));
        private List<ISummaryClause<T>> _summaryElements;

        //ISummaryClause
        public T TrueValue
        {
            set
            {
                _trueValue = value;
                _summaryElements.ForEach(se => se.TrueValue = _trueValue);
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
                _summaryElements.ForEach(se => se.FalseValue = _falseValue);
            }
            get
            {
                return _falseValue;
            }
        }

        //ISummaryClause
        public bool HasCoverage => _summaryElements.Any(se => se.HasCoverage);

        //ISummaryClause
        public bool Covers(T value) => _summaryElements.Any(se => se.Covers(value));

        public bool Empty => !HasCoverage;
        public bool HasExtents => _extents.HasValues;

        private bool ContainsBooleans => typeof(T) == typeof(bool);
        private bool ContainsIntegerNumbers => typeof(T) == typeof(long) || typeof(T) == typeof(Int32) || typeof(T) == typeof(byte);

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

        private void ApplyExtents(ContextExtents<T> extents)
        {
            _extents = extents;
            if (extents.HasValues)
            {
                ApplyExtents(extents.Min, extents.Max);
                return;
            }
        }

        public void ApplyExtents(T min, T max)
        {
            _extents.MinMax(min, max);
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

        public bool HasConditionsNotCoveredBy(ISummaryCoverage summaryCoverage, out ISummaryCoverage diff)
        {
            diff =  CreateSummaryCoverageDifference((SummaryCoverage<T>) summaryCoverage);
            return ((SummaryCoverage<T>)diff).HasCoverage;
        }

        public SummaryCoverage<T> CreateSummaryCoverageDifference(SummaryCoverage<T> toRemove)
        {
            if (toRemove.CoversAllValues || this.Empty)
            {
                return new SummaryCoverage<T>();
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

        public ISummaryCoverage CoverageFor(VBAParser.CaseClauseContext caseClause)
        {
            var caseClauseCoverage = UnreachableSelectCaseFactory.CreateSummaryCoverageShell(this.TypeName);
            foreach (var range in caseClause.rangeClause())
            {
                caseClauseCoverage.Add(CoverageFor(range));
            }
            return caseClauseCoverage;
        }

        public ISummaryCoverage CoverageFor(VBAParser.RangeClauseContext context)
        {
            if (_rangeClauseSummaries.ContainsKey(context))
            {
                return _rangeClauseSummaries[context];
            }
            return UnreachableSelectCaseFactory.CreateSummaryCoverageShell(this.TypeName);
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
            Add((SummaryCoverage<T>)newSummary);
        }

        public void Add(SummaryCoverage<T> newSummary)
        {
            if (!HasExtents && newSummary.HasExtents)
            {
                ApplyExtents(newSummary.Extents.Min, newSummary.Extents.Max);
            }

            if (newSummary.Empty)
            {
                return;
            }

            IsLT.Add(newSummary.IsLT);
            IsGT.Add(newSummary.IsGT);

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

            RelationalOps.Add(newSummary.RelationalOps);
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

        ////Used to modify logic operators to convert LHS and RHS for expressions like '5 > x' (= 'x < 5')
        public static Dictionary<string, string> AlgebraicLogicalInversions = new Dictionary<string, string>()
        {
            [CompareTokens.EQ] = CompareTokens.EQ,
            [CompareTokens.NEQ] = CompareTokens.NEQ,
            [CompareTokens.LT] = CompareTokens.GT,
            [CompareTokens.LTE] = CompareTokens.GTE,
            [CompareTokens.GT] = CompareTokens.LT,
            [CompareTokens.GTE] = CompareTokens.LTE
        };

        public static Dictionary<string, Func<ParseTreeValue, ParseTreeValue, ParseTreeValue>>
            BinaryLogicalOps = new Dictionary<string, Func<ParseTreeValue, ParseTreeValue, ParseTreeValue>>()
            {
                [CompareTokens.GT] = (LHS, RHS) => LHS > RHS ? ParseTreeValue.True : ParseTreeValue.False,
                [CompareTokens.GTE] = (LHS, RHS) => LHS >= RHS ? ParseTreeValue.True : ParseTreeValue.False,
                [CompareTokens.LT] = (LHS, RHS) => LHS < RHS ? ParseTreeValue.True : ParseTreeValue.False,
                [CompareTokens.LTE] = (LHS, RHS) => LHS <= RHS ? ParseTreeValue.True : ParseTreeValue.False,
                [CompareTokens.EQ] = (LHS, RHS) => LHS == RHS ? ParseTreeValue.True : ParseTreeValue.False,
                [CompareTokens.NEQ] = (LHS, RHS) => LHS != RHS ? ParseTreeValue.True : ParseTreeValue.False,
                [Tokens.And] = (LHS, RHS) => LHS.AsBoolean().Value && RHS.AsBoolean().Value ? ParseTreeValue.True : ParseTreeValue.False,
                [Tokens.Or] = (LHS, RHS) => LHS.AsBoolean().Value || RHS.AsBoolean().Value ? ParseTreeValue.True : ParseTreeValue.False,
                [Tokens.XOr] = (LHS, RHS) => LHS.AsBoolean().Value ^ RHS.AsBoolean().Value ? ParseTreeValue.True : ParseTreeValue.False,
                [Tokens.Not] = (LHS, RHS) => LHS.AsBoolean().Value || RHS.AsBoolean().Value ? ParseTreeValue.True : ParseTreeValue.False
                //["Eqv"] = (LHS, RHS) => LHS.AsBoolean().Value ^ RHS.AsBoolean().Value ? UnreachableCaseInspectionValue.True : UnreachableCaseInspectionValue.False
                //["Imp"] = (LHS, RHS) => LHS.AsBoolean().Value ^ RHS.AsBoolean().Value ? UnreachableCaseInspectionValue.True : UnreachableCaseInspectionValue.False,
            };

        private bool RangeStartOrEndHasMismatch(ParserRuleContext prCtxt)
        {
            bool mismatchFound = false;
            foreach (var ctxt in prCtxt.GetChildren<ParserRuleContext>())
            {
                if (ParseTreeValueResults.VariableContexts.Keys.Contains(ctxt))
                {
                    var value = ParseTreeValueResults.VariableContexts[ctxt];
                    if (!mismatchFound)
                    {
                        mismatchFound = !value.HasValueAs(TypeName);
                    }
                }
            }
            return mismatchFound;
        }

        public void LoadRangeClauseCoverage(ParserRuleContext selectStmt, Dictionary<ParserRuleContext,T> valueResolvedContexts)
        {
            TypedValueResults = valueResolvedContexts;

            var rangeClauses = ((VBAParser.SelectCaseStmtContext)selectStmt).caseClause().SelectMany(cc => cc.rangeClause());
            foreach (ParserRuleContext rangeClause in rangeClauses)
            {
                var rangeSummaryCoverage = (SummaryCoverage<T>)UnreachableSelectCaseFactory.CreateSummaryCoverageShell(this.TypeName);
                if (rangeClause.HasChildToken(Tokens.To))
                {
                    var startContext = rangeClause.GetChild<VBAParser.SelectStartValueContext>();
                    var endContext = rangeClause.GetChild<VBAParser.SelectEndValueContext>();

                    var hasStart = TypedValueResults.TryGetValue(startContext, out T startVal);
                    var hasEnd = TypedValueResults.TryGetValue(endContext, out T endVal);

                    if (hasStart && hasEnd)
                    {
                        rangeSummaryCoverage.AddRange(startVal, endVal);
                    }

                    if (!hasStart)
                    {
                        if (RangeStartOrEndHasMismatch(startContext))
                        {
                            IncompatibleTypeRangeContexts.Add(rangeClause);
                        }
                    }
                    if (!hasEnd && hasStart)
                    {
                        if (RangeStartOrEndHasMismatch(endContext))
                        {
                            IncompatibleTypeRangeContexts.Add(rangeClause);
                        }
                    }
                }
                else //single value
                {
                    var ctxts = rangeClause.children.Where(ch => ch is ParserRuleContext
                                    && TypedValueResults.Keys.Contains((ParserRuleContext)ch));

                    //Is Statements
                    if (ctxts.Any() && ctxts.Count() == 1 && rangeClause.HasChildToken(Tokens.Is))
                    {
                        var compOpContext = rangeClause.GetChild<VBAParser.ComparisonOperatorContext>();
                        rangeSummaryCoverage.AddIsClauseResult(compOpContext.GetText(), TypedValueResults[(ParserRuleContext)ctxts.First()]);
                    }
                    //RelationalOp statements like x < 100, 100 < x
                    else if (rangeClause.TryGetChildContext(out VBAParser.RelationalOpContext relOpCtxt))
                    {
                        if (TypedValueResults.Keys.Contains(relOpCtxt))
                        {
                            rangeSummaryCoverage.RelationalOps.Add(TypedValueResults[relOpCtxt]);
                        }
                        else
                        {
                            rangeSummaryCoverage.RelationalOps.Add(relOpCtxt.GetText());
                        }
                    }
                    else if (ctxts.Any() && ctxts.Count() == 1)
                    {
                        rangeSummaryCoverage.Add(TypedValueResults[(ParserRuleContext)ctxts.First()]);
                    }
                }
                _rangeClauseSummaries.Add(rangeClause, rangeSummaryCoverage);
            }
        }

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
            result = $"{IsLT.ToString()}";
            result = IsLT.ToString().Length > 0 ? $"{result}!" : string.Empty;
            result = $"{result}{IsGT.ToString()}";
            result = IsGT.ToString().Length > 0 ? $"{result}!" : string.Empty;
            result = $"{result}{Ranges.ToString()}";
            result = Ranges.ToString().Length > 0 ? $"{result}!" : string.Empty;
            result = $"{result}{SingleValues.ToString()}";
            result = SingleValues.ToString().Length > 0 ? $"{result}!" : string.Empty;
            result = $"{result}{RelationalOps.ToString()}";
            result = RelationalOps.ToString().Length > 0 ? $"{result}!" : string.Empty;
            return result.Length > 0 ? result.Remove(result.Length - 1) : string.Empty;
        }

        private static SummaryCoverage<T> RemoveClausesCoveredBy(SummaryCoverage<T> removeFrom, SummaryCoverage<T> removalSpec)
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

        private static SummaryCoverage<T> RemoveRelationalOpsCoveredBy(SummaryCoverage<T> removeFrom, SummaryCoverage<T> removalSpec)
        {
            if (removalSpec.CoversTrueFalse)
            {
                removeFrom.RelationalOps.Clear();
            }
            return removeFrom;
        }
    }
}
