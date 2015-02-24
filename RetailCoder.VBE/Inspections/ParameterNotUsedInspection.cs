using System.Collections.Generic;
using Rubberduck.VBA.Nodes;

namespace Rubberduck.Inspections
{
    public class ParameterNotUsedInspection : IInspection
    {
        public ParameterNotUsedInspection()
        {
            Severity = CodeInspectionSeverity.Hint;
        }

        public string Name { get { return InspectionNames.ParameterNotUsed; } }
        public CodeInspectionType InspectionType { get { return CodeInspectionType.CodeQualityIssues; } }
        public CodeInspectionSeverity Severity { get; set; }

        public IEnumerable<CodeInspectionResultBase> GetInspectionResults(VBProjectParseResult parseResult)
        {
            var issues = parseResult.IdentifierUsageInspector.UnusedParameters();
            foreach (var issue in issues)
            {
                yield return new ParameterNotUsedInspectionResult(Name, Severity, issue.Context, issue.MemberName);
            }
        }
    }
}