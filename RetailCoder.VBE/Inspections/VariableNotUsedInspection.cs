using System.Collections.Generic;
using System.Linq;
using Rubberduck.VBA.Nodes;

namespace Rubberduck.Inspections
{
    public class VariableNotUsedInspection : IInspection
    {
        public VariableNotUsedInspection()
        {
            Severity = CodeInspectionSeverity.Hint;
        }

        public string Name { get { return InspectionNames.VariableNotUsed; } }
        public CodeInspectionType InspectionType { get { return CodeInspectionType.CodeQualityIssues; } }
        public CodeInspectionSeverity Severity { get; set; }

        public IEnumerable<CodeInspectionResultBase> GetInspectionResults(VBProjectParseResult parseResult)
        {
            var issues = parseResult.IdentifierUsageInspector.AllUnusedVariables();
            foreach (var issue in issues)
            {
                yield return new VariableNotUsedInspectionResult(Name, Severity, issue.Context, issue.QualifiedName);
            }
        }
    }
}