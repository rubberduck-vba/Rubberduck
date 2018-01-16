using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Rubberduck.Inspections.Concrete
{
    public struct SummaryCaseCoverage
    {
        public UnreachableCaseInspectionValue IsLT;
        public UnreachableCaseInspectionValue IsGT;
        public HashSet<UnreachableCaseInspectionValue> SingleValues;
        public List<Tuple<UnreachableCaseInspectionValue, UnreachableCaseInspectionValue>> Ranges;
        public bool CaseElseIsUnreachable;
        public List<string> RangeClausesAsText;
    }

    public class SummaryCoverage
    {
        private SummaryCaseCoverage _summaryCaseCoverage;
        public SummaryCoverage()
        {
            _summaryCaseCoverage = new SummaryCaseCoverage
            {
                IsGT = null,
                IsLT = null,
                SingleValues = new HashSet<UnreachableCaseInspectionValue>(),
                Ranges = new List<Tuple<UnreachableCaseInspectionValue, UnreachableCaseInspectionValue>>(),
                RangeClausesAsText = new List<string>(),
            };
        }

        //public void AddSummary(SummaryCaseCoverage summaryCaseCoverage)
        //{

        //}

        public void AddIsLT(UnreachableCaseInspectionValue isLT)
        {

        }

        public void AddIsGT(UnreachableCaseInspectionValue summaryCaseCoverage)
        {

        }

        public void AddSingleValue(UnreachableCaseInspection singleValue)
        {

        }

        public void AddRange(HashSet<UnreachableCaseInspectionValue> singleValues)
        {

        }

        public void AddRangeClauseText(string rangeClausesAsText)
        {

        }

        public void AddRange(IEnumerable<string> rangeClausesAsText)
        {

        }

        public void AddValueRange(List<Tuple<UnreachableCaseInspectionValue,UnreachableCaseInspectionValue>> valueRange)
        {

        }
    }
}
