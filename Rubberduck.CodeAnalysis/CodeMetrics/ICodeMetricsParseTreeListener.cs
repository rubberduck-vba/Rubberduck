using Antlr4.Runtime.Tree;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor;
using System.Collections.Generic;
using Rubberduck.Parsing.VBA.DeclarationCaching;

namespace Rubberduck.CodeAnalysis.CodeMetrics
{
    public interface ICodeMetricsParseTreeListener : IParseTreeListener
    {
        void Reset();
        void InjectContext((DeclarationFinder, QualifiedModuleName) context);
        IEnumerable<ICodeMetricResult> Results();
    }
}
