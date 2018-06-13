using Antlr4.Runtime.Misc;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using System.Collections.Generic;

namespace Rubberduck.CodeAnalysis.CodeMetrics
{
    internal class LineCountModuleMetric : CodeMetric
    {
        public LineCountModuleMetric() : base("Line Count", AggregationLevel.Module)
        {
            TreeListener = new LineCountModuleListener(this);
        }
        public override ICodeMetricsParseTreeListener TreeListener { get; }
    }

    internal class LineCountModuleMetricResult : CodeMetricResultBase
    {
        private readonly int value;

        public LineCountModuleMetricResult(Declaration declaration, CodeMetric metricReference, int value)
            : base(declaration, metricReference)
        {
            this.value = value;
        }
        public override string Value => value.ToString();
    }

    internal class LineCountModuleListener : CodeMetricListenerBase
    {
        private int workingValue;

        public LineCountModuleListener(CodeMetric owner) : base(owner) { }

        public override void Reset()
        {
            base.Reset();
            workingValue = 0;
        }

        public override IEnumerable<ICodeMetricResult> Results() => new[] { new LineCountModuleMetricResult(_finder.ModuleDeclaration(_qmn), ownerReference, workingValue) };

        public override void EnterEndOfLine([NotNull]  VBAParser.EndOfLineContext _) => workingValue++;
    }
}
