using System.Collections.Generic;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA.DeclarationCaching;
using Rubberduck.VBEditor;

namespace Rubberduck.CodeAnalysis.CodeMetrics
{
    internal abstract class CodeMetricResultBase : ICodeMetricResult
    {
        public CodeMetricResultBase(Declaration declaration, CodeMetric metric)
        {
            Declaration = declaration;
            Metric = metric;
        }

        public Declaration Declaration { get; }

        public CodeMetric Metric { get; }

        public abstract string Value { get; }
    }

    internal abstract class CodeMetricListenerBase : VBAParserBaseListener, ICodeMetricsParseTreeListener
    {
        protected readonly CodeMetric ownerReference;
        protected DeclarationFinder _finder;
        protected QualifiedModuleName _qmn;

        public CodeMetricListenerBase(CodeMetric metric)
        {
            ownerReference = metric;
        }

        public void InjectContext((DeclarationFinder, QualifiedModuleName) context) => (_finder, _qmn) = context;
        public virtual void Reset()
        {
            _finder = null;
            _qmn = default;
        }
        public abstract IEnumerable<ICodeMetricResult> Results();
    }
}
