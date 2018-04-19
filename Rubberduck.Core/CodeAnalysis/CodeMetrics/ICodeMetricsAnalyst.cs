using Rubberduck.Parsing.VBA;
using System.Collections.Generic;

namespace Rubberduck.CodeAnalysis.CodeMetrics
{
    public interface ICodeMetricsAnalyst
    {
        IEnumerable<ICodeMetricResult> GetMetrics(RubberduckParserState state);
    }
}
