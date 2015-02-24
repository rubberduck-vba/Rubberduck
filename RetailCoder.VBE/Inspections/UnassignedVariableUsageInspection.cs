using System.Collections.Generic;
using Rubberduck.VBA.Nodes;

namespace Rubberduck.Inspections
{
    public class UnassignedVariableUsageInspection : IInspection
    {
        public UnassignedVariableUsageInspection()
        {
            Severity = CodeInspectionSeverity.Error;
        }

        public string Name { get { return InspectionNames.UnassignedVariableUsage; } }
        public CodeInspectionType InspectionType { get { return CodeInspectionType.CodeQualityIssues; } }
        public CodeInspectionSeverity Severity { get; set; }

        public IEnumerable<CodeInspectionResultBase> GetInspectionResults(VBProjectParseResult parseResult)
        {
            var issues = parseResult.IdentifierUsageInspector.AllUnassignedVariableUsages();
            foreach (var issue in issues)
            {
                yield return new UnassignedVariableUsageInspectionResult(Name, Severity, issue.Context, issue.QualifiedName);
            }
        }
    }
}