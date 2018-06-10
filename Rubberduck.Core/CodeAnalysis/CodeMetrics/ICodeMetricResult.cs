using Rubberduck.Parsing.Symbols;

namespace Rubberduck.CodeAnalysis.CodeMetrics
{
    /// <summary>
    /// A CodeMetricsResult. Each result is attached to a Declaration.
    /// Usually this declaration would be a Procedure (Function/Sub/Property).
    /// Some metrics are only useful on Module level, some even on Project level.
    /// 
    /// Some metrics may be aggregated to obtain a metric for a "higher hierarchy level"
    /// </summary>
    public interface ICodeMetricResult
    {
        /// <summary>
        /// The declaration that this result refers to.
        /// </summary>
        Declaration Declaration { get; }
        /// <summary>
        /// The Metric kind that this result belongs to. Only results belonging to the **same** metric can be aggregated.
        /// </summary>
        CodeMetric Metric { get; }
        /// <summary>
        /// A string representation of the value.
        /// </summary>
        string Value { get; }
    }
}
