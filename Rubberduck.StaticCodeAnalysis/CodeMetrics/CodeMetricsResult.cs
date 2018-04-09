using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor;
using System;
using System.Collections.Generic;
using System.Linq;

namespace Rubberduck.CodeAnalysis.CodeMetrics
{
    public struct CodeMetricsResult : ICodeMetricsResult
    {
        public CodeMetricsResult(int lines, int cyclomaticComplexity, int nesting)
            : this(lines, cyclomaticComplexity, nesting, Enumerable.Empty<ICodeMetricsResult>())
        { 
        }

        public CodeMetricsResult(int lines, int cyclomaticComplexity, int nesting, IEnumerable<ICodeMetricsResult> childScopeResults)
        {
            var childScopeMetric =
                childScopeResults.Aggregate(new CodeMetricsResult(), (r1, r2) => new CodeMetricsResult(r1.Lines + r2.Lines, r1.CyclomaticComplexity + r2.CyclomaticComplexity, Math.Max(r1.MaximumNesting, r2.MaximumNesting)));
            Lines = lines + childScopeMetric.Lines;
            CyclomaticComplexity = cyclomaticComplexity + childScopeMetric.CyclomaticComplexity;
            MaximumNesting = Math.Max(nesting, childScopeMetric.MaximumNesting);
        }
        
        public int Lines { get; private set; }
        public int CyclomaticComplexity { get; private set; }
        public int MaximumNesting { get; private set; }

    }

    public struct MemberMetricsResult : IMemberMetricsResult
    {
        public Declaration Member { get; private set; }
        public ICodeMetricsResult Result { get; private set; }

        public MemberMetricsResult(Declaration member, IEnumerable<ICodeMetricsResult> contextResults)
        {
            Member = member;
            Result = new CodeMetricsResult(0, 0, 0, contextResults);
        }
    }

    public struct ModuleMetricsResult : IModuleMetricsResult
    {
        public QualifiedModuleName ModuleName { get; private set; }
        public ICodeMetricsResult Result { get; private set; }
        public IReadOnlyDictionary<Declaration, ICodeMetricsResult> MemberResults { get; private set; }

        public ModuleMetricsResult(QualifiedModuleName moduleName, IEnumerable<IMemberMetricsResult> memberMetricsResult, IEnumerable<ICodeMetricsResult> nonMemberResults)
        {
            ModuleName = moduleName;
            MemberResults = memberMetricsResult.ToDictionary(mmr => mmr.Member, mmr => mmr.Result);
            Result = new CodeMetricsResult(0, 0, 0, nonMemberResults.Concat(MemberResults.Values));
        }
    }
}
