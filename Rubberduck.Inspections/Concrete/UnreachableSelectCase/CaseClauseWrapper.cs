using Rubberduck.Parsing.Grammar;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Rubberduck.Inspections.Concrete
{
    //public class CaseClauseWrapper<T> where T : IComparable<T>
    //{
    //    ContextValueResults<T> _ctxtValueResults;
    //    IParseTreeValueResults _parseTreeResults;
    //    private readonly VBAParser.CaseClauseContext _caseClause;

    //    private List<RangeClauseWrapper<T>> _rangeClauseWrappers;
    //    private SummaryCoverage<T> _summaryCoverage;

    //    public CaseClauseWrapper(VBAParser.CaseClauseContext caseClause, ContextValueResults<T> ctxtValueResults)
    //    {
    //        _ctxtValueResults = ctxtValueResults;
    //        _caseClause = caseClause;
    //        _rangeClauseWrappers = new List<RangeClauseWrapper<T>>();
    //        _summaryCoverage = LoadSummaryCoverage(new SummaryCoverage<T>(_ctxtValueResults.Extents, _ctxtValueResults.TrueValue, _ctxtValueResults.FalseValue));
    //    }

    //    public CaseClauseWrapper(VBAParser.CaseClauseContext caseClause, IParseTreeValueResults ptResults)
    //    {
    //        //_ctxtValueResults = ctxtValueResults;
    //        _parseTreeResults = ptResults;
    //        _caseClause = caseClause;
    //        _rangeClauseWrappers = new List<RangeClauseWrapper<T>>();
    //        _summaryCoverage = LoadSummaryCoverage(new SummaryCoverage<T>(_ctxtValueResults.Extents, _ctxtValueResults.TrueValue, _ctxtValueResults.FalseValue));
    //    }
        

    //    public SummaryCoverage<T> SummaryCoverage => _summaryCoverage;

    //    public bool CanBeInspected => (_summaryCoverage.HasCoverage || _summaryCoverage.HasExtents)
    //                && _rangeClauseWrappers.Any(rcw => rcw.CanBeInspected)
    //                && !IsIncompatibleType;

    //    public bool IsIncompatibleType => _rangeClauseWrappers.All(rcw => rcw.IsIncompatibleType);

    //    public bool HasConditionsNotCoveredBy(SummaryCoverage<T> existingSummaryCoverage, out SummaryCoverage<T> difference)
    //    {
    //        difference = SummaryCoverage.CreateSummaryCoverageDifference(existingSummaryCoverage);
    //        return difference.HasCoverage;
    //    }

    //    public SummaryCoverage<T> LoadSummaryCoverage(SummaryCoverage<T> caseClauseSummaryCoverage)
    //    {
    //        var rgClauses = _caseClause.rangeClause();
    //        foreach (var rgClause in rgClauses)
    //        {
    //            var wrapper = new RangeClauseWrapper<T>(rgClause, _ctxtValueResults);
    //            caseClauseSummaryCoverage.Add(wrapper.SummaryCoverage);
    //            _rangeClauseWrappers.Add(wrapper);
    //        }
    //        return caseClauseSummaryCoverage;
    //    }
    //}
}
