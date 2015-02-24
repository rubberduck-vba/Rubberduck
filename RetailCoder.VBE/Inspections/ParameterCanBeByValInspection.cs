using System.Collections.Generic;
using Rubberduck.VBA.Nodes;

namespace Rubberduck.Inspections
{
    public class ParameterCanBeByValInspection : IInspection
    {
        public ParameterCanBeByValInspection()
        {
            Severity = CodeInspectionSeverity.Suggestion;
        }

        public string Name { get { return InspectionNames.ParameterCanBeByVal; } }
        public CodeInspectionType InspectionType { get { return CodeInspectionType.CodeQualityIssues; } }
        public CodeInspectionSeverity Severity { get; set; }

        public IEnumerable<CodeInspectionResultBase> GetInspectionResults(VBProjectParseResult parseResult)
        {
            var issues = parseResult.IdentifierUsageInspector.UnassignedByRefParameters();

            foreach (var issue in issues)
            {
                yield return new ParameterCanBeByValInspectionResult(Name, Severity, issue.Context, issue.MemberName);
            }
        }
    }
}