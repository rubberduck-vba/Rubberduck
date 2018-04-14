using Antlr4.Runtime.Tree;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor;
using System;
using System.Collections.Generic;

namespace Rubberduck.CodeAnalysis.CodeMetrics
{
    public interface ICodeMetricsAnalyst
    {
        IEnumerable<ICodeMetricResult> GetMetrics(RubberduckParserState state);
    }

    public abstract class CodeMetric
    {
        public CodeMetric(string name, AggregationLevel level) => (Name, Level) = (name, level);

        /// <summary>
        /// The name of the metric. Used for localization purposes as well as a uniquely identifying name to disambiguate between metrics.
        /// </summary>
        public string Name { get; }
        /// <summary>
        /// The aggregation level (or levels) that this metric applies to. If multiple levels apply, instances of ICodeMetricResult can be aggregated.
        /// </summary>
        public AggregationLevel Level { get; }

        public abstract ICodeMetricsParseTreeListener TreeListener { get;  }
    }

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

    [Flags]
    public enum AggregationLevel
    {
        Project = 1 << 0,
        Module = 1 << 1,
        Member = 1 << 2,
        Declaration = 1 << 3,
    }

    public interface ICodeMetricsParseTreeListener : IParseTreeListener
    {
        void Reset();
        void InjectContext((DeclarationFinder, QualifiedModuleName) context);
        IEnumerable<ICodeMetricResult> Results();
    }
}
