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

    internal static class MathTokens
    {
        public static readonly string MULT = "*";
        public static readonly string DIV = "/";
        public static readonly string ADD = "+";
        public static readonly string SUBTRACT = "-";
        public static readonly string POW = "^";
        public static readonly string MOD = Tokens.Mod;
    }

    public class SummaryCoverage<T> where T : System.IComparable<T>
    {
        private ContextExtents<T> _extents;
        private bool _coversAllValues;

        public SummaryCoverage()
        {
            _coversAllValues = false;
            ApplyExtents(new ContextExtents<T>());
        }

        public SummaryCoverage(ContextExtents<T> extents)
        {
            _coversAllValues = false;
            ApplyExtents(extents);
        }

        private ContextExtents<T> Extents => _extents;
        public SummaryClauseRanges<T> Ranges { set; get; } = new SummaryClauseRanges<T>();
        public SummaryClauseSingleValues<T> SingleValues { set; get; } = new SummaryClauseSingleValues<T>();
        public SummaryClauseIsLT<T> IsLT { set; get; } = new SummaryClauseIsLT<T>();
        public SummaryClauseIsGT<T> IsGT { set; get; } = new SummaryClauseIsGT<T>();
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

        public bool HasCoverage => SingleValues.Any()  || Ranges.Any()
            || IsLT.HasCoverage || IsGT.HasCoverage;

        public bool Empty => !HasCoverage;
        public bool HasExtents => _extents.HasValues;

        private bool ContainsBooleans => typeof(T) == typeof(bool);
        private bool ContainsIntegerNumbers => typeof(T) == typeof(long) || typeof(T) == typeof(Int32) || typeof(T) == typeof(byte);

        private void ApplyExtents( ContextExtents<T> extents)
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

            return ltIsEqual && gtIsEqual && rangesAreEqual && singleValuesAreEqual; // && boolsAreEqual;
        }

        public override int GetHashCode()
        {
            var hashString = IsLT.ToString();
            hashString = hashString + IsGT.ToString();
            hashString = hashString + Ranges.ToString();
            hashString = hashString + SingleValues.ToString();
            return hashString.GetHashCode();
        }

        public SummaryCoverage<T> RemoveCoverageRedundantTo(SummaryCoverage<T> toRemove)
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
            if (!ContainsBooleans)
            {
                IsLT.Value = newVal;
            }
        }

        public void SetIsGT(T newVal)
        {
            if (!ContainsBooleans)
            {
                IsGT.Value = newVal;
            }
        }

        public void AddRange(T start, T end)
        {
            var candidate = new SummaryClauseRange<T>(start, end);
            AddRange(candidate);
        }

        private void AddRange(SummaryClauseRange<T> range)
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

        private void Add(bool value)
        {
            SingleValues.Add(value);
        }

        public void Add(T value)
        {
            if(!(IsLT.Covers(value) || IsGT.Covers(value)))
            {
                SingleValues.Add(value);
            }
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

            foreach (var val in newSummary.SingleValues.Values)
            {
                Add(val);
            }

            foreach (var val in newSummary.SingleValues.ValuesBoolean)
            {
                Add(val);
            }

            //if (!ContainsBooleans)
            //{
            //    if (newSummary.BooleanValues.Any())
            //    {
            //        foreach (var boolVal in newSummary.BooleanValues)
            //        {
            //            BooleanValues.Add(boolVal);
            //        }
            //    }
            //}
            //else
            {
                IsLT.Add(newSummary.IsLT);
                IsGT.Add(newSummary.IsGT);
            }

            foreach (var range in newSummary.Ranges.RangeClauses)
            {
                //if (ContainsBooleans)
                //{
                //    SingleValues.Add(range.Start);
                //    SingleValues.Add(range.End);
                //}
                //else
                {
                    AddRange(range);
                }
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

        //Used to modify logic operators to convert LHS and RHS for expressions like '5 > x' (= 'x < 5')
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
            BinaryMathOps = new Dictionary<string, Func<ParseTreeValue, ParseTreeValue, ParseTreeValue>>()
            {
                [MathTokens.ADD] = (LHS, RHS) => LHS + RHS,
                [MathTokens.SUBTRACT] = (LHS, RHS) => LHS - RHS,
                [MathTokens.MULT] = (LHS, RHS) => LHS * RHS,
                [MathTokens.DIV] = (LHS, RHS) => LHS / RHS,
                [MathTokens.POW] = (LHS, RHS) => ParseTreeValue.Pow(LHS, RHS),
                [MathTokens.MOD] = (LHS, RHS) => LHS % RHS
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

        private static Dictionary<string, Func<ParseTreeValue, ParseTreeValue>>
            UnaryLogicalOps = new Dictionary<string, Func<ParseTreeValue, ParseTreeValue>>()
            {
                [Tokens.Not] = (LHS) => !(LHS.AsBoolean().Value) ? ParseTreeValue.True : ParseTreeValue.False
            };

        public void LoadCoverage(VBAParser.CaseClauseContext caseClause, ContextValueResults<T> ctxtValueResults)
        {
            if (ctxtValueResults.Extents.HasValues)
            {
                ApplyExtents(ctxtValueResults.Extents.Min, ctxtValueResults.Extents.Max);
            }

            var rgClauses = caseClause.children.Where(ch => ch is VBAParser.RangeClauseContext);
            foreach (ParserRuleContext rangeClause in rgClauses)
            {
                if (rangeClause.HasChildToken(Tokens.To))
                {
                    var startContext = rangeClause.GetChild<VBAParser.SelectStartValueContext>();
                    var endContext = rangeClause.GetChild<VBAParser.SelectEndValueContext>();
                    if (ctxtValueResults.ValueResolvedContexts.TryGetValue(startContext, out T startVal) &&
                            ctxtValueResults.ValueResolvedContexts.TryGetValue(endContext, out T endVal))
                    {
                        AddRange(startVal, endVal);
                    }
                }
                else //single value
                {
                    var ctxts = rangeClause.children.Where(ch => ch is ParserRuleContext
                                    && ctxtValueResults.ValueResolvedContexts.Keys.Contains((ParserRuleContext)ch));

                    //Is Statements
                    if (ctxts.Any() && ctxts.Count() == 1 && rangeClause.HasChildToken(Tokens.Is))
                    {
                        var compOpContext = rangeClause.GetChild<VBAParser.ComparisonOperatorContext>();
                        AddIsClauseResult(compOpContext.GetText(), ctxtValueResults.ValueResolvedContexts[(ParserRuleContext)ctxts.First()]);
                    }
                    //RelationalOp statements like x < 100, 100 < x
                    else if (rangeClause.TryGetChildContext(out VBAParser.RelationalOpContext relOpCtxt))
                    {
                        if (!ctxtValueResults.ValueResolvedContexts.Keys.Contains(relOpCtxt))
                        {
                            var relOpContexts = relOpCtxt.children.Where(ch => !(ch is VBAParser.WhiteSpaceContext)).ToList();
                            for (var idx = 0; idx < relOpContexts.Count(); idx++)
                            {
                                var ctxt = relOpContexts[idx];
                                if (ctxtValueResults.ValueResolvedContexts.Keys.Contains(ctxt) )
                                {
                                    var opSymbol = relOpCtxt.children.Where(ch => BinaryLogicalOps.Keys.Contains(ch.GetText())).First().GetText();
                                    if (idx == 0)
                                    {
                                        //100 < x: when the value is the first child, the expression's opSymbol
                                        //needs to be converted to represent x < 100
                                        AddIsClauseResult(AlgebraicLogicalInversions[opSymbol], ctxtValueResults.ValueResolvedContexts[(ParserRuleContext)ctxt]);
                                    }
                                    else
                                    {
                                        AddIsClauseResult(opSymbol, ctxtValueResults.ValueResolvedContexts[(ParserRuleContext)ctxt]);
                                    }
                                }
                            }
                        }
                    }
                    else if (ctxts.Any() && ctxts.Count() == 1)
                    {
                        Add(ctxtValueResults.ValueResolvedContexts[(ParserRuleContext)ctxts.First()]);
                    }
                }
            }
        }

        public void AddIsClauseResult(string compareOperator, T result)
        {
            if (compareOperator.Equals(CompareTokens.LT))
            {
                SetIsLT(result);

            }
            else if (compareOperator.Equals(CompareTokens.LTE))
            {
                SetIsLT(result);
                Add(result);
            }
            else if (compareOperator.Equals(CompareTokens.GT))
            {
                SetIsGT(result);
            }
            else if (compareOperator.Equals(CompareTokens.GTE))
            {
                SetIsGT(result);
                Add(result);
            }
            else if (compareOperator.Equals(CompareTokens.EQ))
            {
                Add(result);
            }
            else if (compareOperator.Equals(CompareTokens.NEQ))
            {
                if (ContainsBooleans)
                {
                    SingleValues.Add(!result.Equals(true));
                }
                SetIsLT(result);
                SetIsGT(result);
            }
            else
            {
                Debug.Assert(false, "Unrecognized comparison symbol for Is Clause");
            }
        }

        public override string ToString()
        {
            var result = string.Empty;
            result = $"{result}{IsLT.ToString()}";
            result = IsLT.ToString().Length > 0 ? $"{result}," : string.Empty;
            result = $"{result}{IsGT.ToString()}";
            result = IsGT.ToString().Length > 0 ? $"{result}," : string.Empty;
            result = $"{result}{Ranges.ToString()}";
            result = Ranges.ToString().Length > 0 ? $"{result}," : string.Empty;
            result = $"{result}{SingleValues.ToString()}";
            result = SingleValues.ToString().Length > 0 ? $"{result}," : string.Empty;
            return result.Length > 0 ? result.Remove(result.Length - 1) : string.Empty;
        }

        private static SummaryCoverage<T> RemoveClausesCoveredBy(SummaryCoverage<T> removeFrom, SummaryCoverage<T> removalSpec)
        {
            var newSummary = RemoveIsClausesCoveredBy(removeFrom, removalSpec);
            newSummary = RemoveRangesCoveredBy(removeFrom, removalSpec);
            return RemoveSingleValuesCoveredBy(removeFrom, removalSpec);
        }

        private static SummaryCoverage<T> RemoveIsClausesCoveredBy(SummaryCoverage<T> removeFrom, SummaryCoverage<T> removalSpec)
        {
            if (!removeFrom.ContainsBooleans)
            {
                if (removalSpec.IsLT.HasCoverage && removalSpec.IsLT.Covers(removeFrom.IsLT))
                {
                    removeFrom.IsLT.Reset();
                }
                if (removalSpec.IsGT.HasCoverage && removalSpec.IsGT.Covers(removeFrom.IsGT))
                {
                    removeFrom.IsGT.Reset();
                }
            }
            return removeFrom;
        }

        private static SummaryCoverage<T> RemoveRangesCoveredBy(SummaryCoverage<T> removeFrom, SummaryCoverage<T> removalSpec)
        {
            var toRemove = removeFrom.Ranges.RangeClauses.Where(rg => removalSpec.IsLT.Covers(rg.Start) && removalSpec.IsLT.Covers(rg.End)
                    || removalSpec.IsGT.Covers(rg.Start) && removalSpec.IsGT.Covers(rg.End)).ToList();

            for (var idx = 0; idx < removeFrom.Ranges.RangeClauses.Count; idx++)
            {
                var rangeClause = removeFrom.Ranges.RangeClauses[idx];
                if (removalSpec.Ranges.RangeClauses.Any(rg => rg.Covers(rangeClause)))
                {
                    toRemove.Add(rangeClause);
                }
            }

            removeFrom.Ranges.Remove(toRemove);
            return removeFrom;
        }

        private static SummaryCoverage<T> RemoveSingleValuesCoveredBy(SummaryCoverage<T> removeFrom, SummaryCoverage<T> removalSpec)
        {
            List<T> toRemove = new List<T>();
            List<bool> toRemoveBools = new List<bool>();
            toRemove = removeFrom.SingleValues.Values.Where(sv => removalSpec.SingleValues.Covers(sv)).ToList();
            toRemoveBools = removeFrom.SingleValues.ValuesBoolean.Where(sv => removalSpec.SingleValues.Covers(sv)).ToList();
            toRemove.AddRange(removeFrom.SingleValues.Values.Where(sv => removalSpec.IsLT.Covers(sv) || removalSpec.IsGT.Covers(sv)).ToList());
            foreach (var range in removalSpec.Ranges.RangeClauses)
            {
                toRemove.AddRange(removeFrom.SingleValues.Values.Where(sv => range.Covers(sv)));
            }
            toRemove.AddRange(removeFrom.SingleValues.Values.Where(sv => removalSpec.SingleValues.Covers(sv)));

            removeFrom.SingleValues.Remove(toRemove);
            removeFrom.SingleValues.Remove(toRemoveBools);
            return removeFrom;
        }
    }
}
