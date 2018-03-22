using Antlr4.Runtime;
using Rubberduck.Parsing.Grammar;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Rubberduck.Inspections.Concrete.UnreachableSelectCase
{
    public class SummaryClauseRelationalOps<T> : SummaryClauseSingleValueBase<T> where T : IComparable<T>
    {
        private ISummaryClauseSingleValues<T> _singleValues;
        private List<string> _variableRelationalOps;

        public SummaryClauseRelationalOps(ISummaryClauseSingleValues<T> singleValues) : base(singleValues.TConverter)
        {
            _singleValues = singleValues;
            _variableRelationalOps = new List<string>();
        }

        public override bool HasCoverage => _variableRelationalOps.Any();
        public override bool Covers(T candidate) => _singleValues.Covers(candidate);
        //public bool Covers(string relationalOpText) => _unresolvedRelationalOps.Contains(relationalOpText);

        public void Clear()
        {
            _variableRelationalOps.Clear();
        }

        public override void Add(T value)
        {
            if (!(Covers(TrueValue) && Covers(FalseValue)))
            {
                if (value.CompareTo(FalseValue) != 0)
                {
                    _singleValues.Add(TrueValue);
                }
                else
                {
                    _singleValues.Add(FalseValue);
                }
            }
        }

        public void Add(SummaryClauseRelationalOps<T> newRelOp)
        {
            if (newRelOp.HasCoverage && !(Covers(TrueValue) && Covers(FalseValue)))
            {
                _variableRelationalOps.AddRange(newRelOp._variableRelationalOps);
            }
        }

        public void Add(string value)
        {
            if(!(Covers(TrueValue) && Covers(FalseValue)))
            {
                if (!_variableRelationalOps.Contains(value))
                {
                    _variableRelationalOps.Add(value);
                }
            }
        }

        public int Count => _variableRelationalOps.Count;

        public override string ToString()
        {
            if (!_variableRelationalOps.Any())
            {
                return string.Empty;
            }
            const string prefix = "RelOp=";
            var result = prefix;
            foreach (var val in _variableRelationalOps)
            {
                result = $"{result}{val.ToString()},";
            }
            return result.Length > 0 ? result.Remove(result.Length - 1) : string.Empty;
        }
    }
}
