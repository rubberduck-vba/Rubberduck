namespace Rubberduck.CodeAnalysis.CodeMetrics
{
    public abstract class CodeMetric
    {
        public CodeMetric(string name, AggregationLevel level) => (Name, Level) = (name, level);

        /// <summary>
        /// The name of the metric. Used for localization purposes as well as a uniquely identifying name to disambiguate between metrics.
        /// </summary>
        public string Name { get; }
        /// <summary>
        /// The aggregation level that this metric applies to.
        /// </summary>
        public AggregationLevel Level { get; }

        public abstract ICodeMetricsParseTreeListener TreeListener { get;  }
    }
}
