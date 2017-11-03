using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor;
using System;
using System.Collections.Generic;
using System.Linq;

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
        
        public int Lines { get; private set; }
        public int CyclomaticComplexity { get; private set; }
        public int MaximumNesting { get; private set; }

    }

    public struct MemberMetricsResult
    {
        public Declaration Member { get; private set; }
        public CodeMetricsResult Result { get; private set; }

        public MemberMetricsResult(Declaration member, IEnumerable<CodeMetricsResult> contextResults)
        {
            Member = member;
            Result = new CodeMetricsResult(0, 0, 0, contextResults);
        }
    }

    public struct ModuleMetricsResult
    {
        public QualifiedModuleName ModuleName { get; private set; }
        public CodeMetricsResult Result { get; private set; }
        public IReadOnlyDictionary<Declaration, CodeMetricsResult> MemberResults { get; private set; }

        public ModuleMetricsResult(QualifiedModuleName moduleName, IEnumerable<MemberMetricsResult> memberMetricsResult, IEnumerable<CodeMetricsResult> nonMemberResults)
        {
            ModuleName = moduleName;
            MemberResults = memberMetricsResult.ToDictionary(mmr => mmr.Member, mmr => mmr.Result);
            Result = new CodeMetricsResult(0, 0, 0, nonMemberResults.Concat(MemberResults.Values));
        }
    }
}
