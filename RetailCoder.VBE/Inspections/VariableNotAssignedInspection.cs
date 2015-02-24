using System.Collections.Generic;
using Rubberduck.VBA.Nodes;

namespace Rubberduck.Inspections
{
    public class VariableNotAssignedInspection : IInspection
    {
        public VariableNotAssignedInspection()
        {
            Severity = CodeInspectionSeverity.Error;
        }

        public string Name { get { return InspectionNames.VariableNotAssigned; } }
        public CodeInspectionType InspectionType { get { return CodeInspectionType.CodeQualityIssues; } }
        public CodeInspectionSeverity Severity { get; set; }

        public IEnumerable<CodeInspectionResultBase> GetInspectionResults(VBProjectParseResult parseResult)
        {
            var issues = parseResult.IdentifierUsageInspector.AllUnassignedVariables();
            foreach (var issue in issues)
            {
                yield return new VariableNotAssignedInspectionResult(Name, Severity, issue.Context, issue.QualifiedName);
            }
        }
    }
}