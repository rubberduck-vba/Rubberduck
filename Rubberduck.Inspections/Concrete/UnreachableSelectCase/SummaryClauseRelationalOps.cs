using Antlr4.Runtime;
using Rubberduck.Parsing.Grammar;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Rubberduck.Inspections.Concrete
{
    public class SummaryClauseRelationalOps<T> : SummaryClauseSingleValueBase<T> where T : IComparable<T>
    {
        private ISummaryClauseSingleValues<T> _singleValues;
        private List<string> _unresolvedRelationalOps;

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

        public SummaryClauseRelationalOps(ISummaryClauseSingleValues<T> singleValues)
        {
            _singleValues = singleValues;
            _unresolvedRelationalOps = new List<string>();
        }

        public override bool HasCoverage => _unresolvedRelationalOps.Any();
        public override bool Covers(T candidate) => _singleValues.Covers(candidate);
        public bool Covers(string relationalOpText) => _unresolvedRelationalOps.Contains(relationalOpText);
        public override void Add(T value)
        {
            if (!(Covers(TrueValue) && Covers(FalseValue)))
            {
                if(value.CompareTo(FalseValue) != 0)
                {
                    _singleValues.Add(TrueValue);
                }
                else
                {
                    _singleValues.Add(FalseValue);
                }
            }
        }

        public void Clear()
        {
            _unresolvedRelationalOps.Clear();
        }

        public void Add(SummaryClauseRelationalOps<T> newRelOp)
        {
            if (newRelOp.HasCoverage && !(Covers(TrueValue) && Covers(FalseValue)))
            {
                _unresolvedRelationalOps.AddRange(newRelOp._unresolvedRelationalOps);
            }
        }

        public void Add(string value)
        {
            if(!(Covers(TrueValue) && Covers(FalseValue)))
            {
                if (!_unresolvedRelationalOps.Contains(value))
                {
                    _unresolvedRelationalOps.Add(value);
                }
            }
        }

        public int Count => _unresolvedRelationalOps.Count;

        public override string ToString()
        {
            var result = string.Empty;
            foreach (var val in _unresolvedRelationalOps)
            {
                result = $"{result}RelOp={val},";
            }
            return result.Length > 0 ? result.Remove(result.Length - 1) : string.Empty;

        }
    }
}
