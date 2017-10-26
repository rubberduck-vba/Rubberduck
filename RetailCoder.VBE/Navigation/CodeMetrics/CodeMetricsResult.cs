using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Rubberduck.Navigation.CodeMetrics
{
    public struct CodeMetricsResult
    {
        public CodeMetricsResult(int lines, int cyclomaticComplexity, int nesting)
            : this(lines, cyclomaticComplexity, nesting, Enumerable.Empty<CodeMetricsResult>())
        { 
        }

        public CodeMetricsResult(int lines, int cyclomaticComplexity, int nesting, IEnumerable<CodeMetricsResult> childScopeResults)
        {
            var childScopeMetric =
                childScopeResults.Aggregate(new CodeMetricsResult(), (r1, r2) => new CodeMetricsResult(r1.Lines + r2.Lines, r1.CyclomaticComplexity + r2.CyclomaticComplexity, Math.Max(r1.MaximumNesting, r2.MaximumNesting)));
            Lines = lines + childScopeMetric.Lines;
            CyclomaticComplexity = cyclomaticComplexity + childScopeMetric.CyclomaticComplexity;
            MaximumNesting = Math.Max(nesting, childScopeMetric.MaximumNesting);
        }

        // possibly refer to a selection?
        public int Lines { get; private set; }
        public int CyclomaticComplexity { get; private set; }
        public int MaximumNesting { get; private set; }

    }
}
