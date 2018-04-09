using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor;
using System.Collections.Generic;

namespace Rubberduck.CodeAnalysis.CodeMetrics
{
    public interface ICodeMetricsResult
    {
        int CyclomaticComplexity { get; }
        int Lines { get; }
        int MaximumNesting { get; }
    }

    public interface IModuleMetricsResult
    {
        IReadOnlyDictionary<Declaration, ICodeMetricsResult> MemberResults { get; }
        QualifiedModuleName ModuleName { get; }
        ICodeMetricsResult Result { get; }
    }

    public interface IMemberMetricsResult
    {
        Declaration Member { get; }
        ICodeMetricsResult Result { get; }
    }
}