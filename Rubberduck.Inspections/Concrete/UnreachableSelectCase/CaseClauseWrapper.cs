using Rubberduck.Parsing.Grammar;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Rubberduck.Inspections.Concrete
{
    public class CaseClauseWrapper<T> where T : IComparable<T>
    {
        ContextValueResults<T> _ctxtValueResults;
        private readonly VBAParser.CaseClauseContext _caseClause;

        private List<RangeClauseWrapper<T>> _rangeClauseWrappers;
        private SummaryCoverage<T> _summaryCoverage;

        public CaseClauseWrapper(VBAParser.CaseClauseContext caseClause, ContextValueResults<T> ctxtValueResults)
        {
            _ctxtValueResults = ctxtValueResults;
            _caseClause = caseClause;
            _rangeClauseWrappers = new List<RangeClauseWrapper<T>>();
            _summaryCoverage = LoadSummaryCoverage(new SummaryCoverage<T>(_ctxtValueResults.Extents));
        }

        public SummaryCoverage<T> SummaryCoverage => _summaryCoverage;
        public bool CanBeInspected => (SummaryCoverage.HasCoverage || SummaryCoverage.HasExtents)
                    && _rangeClauseWrappers.Any(rcw => rcw.CanBeInspected)
                    && !IsMismatch;

        public bool IsMismatch => _rangeClauseWrappers.All(rcw => rcw.IsMismatch);

        public SummaryCoverage<T> RemoveCoverageRedundantTo(SummaryCoverage<T> summaryCoverage)
        {
            return SummaryCoverage.RemoveCoverageRedundantTo(summaryCoverage);
        }

        private SummaryCoverage<T> LoadSummaryCoverage(SummaryCoverage<T> caseClauseSummaryCoverage)
        {
            var rgClauses = _caseClause.rangeClause();
            foreach (var rgClause in rgClauses)
            {
                var wrapper = new RangeClauseWrapper<T>(rgClause, _ctxtValueResults);
                caseClauseSummaryCoverage.Add(wrapper.SummaryCoverage);
                _rangeClauseWrappers.Add(wrapper);
            }
            return caseClauseSummaryCoverage;
        }
    }
}
