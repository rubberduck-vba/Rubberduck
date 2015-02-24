using System.Collections.Generic;
using Rubberduck.VBA.Nodes;

namespace Rubberduck.Inspections
{
    public class VariableNotDeclaredInspection : IInspection
    {
        public VariableNotDeclaredInspection()
        {
            Severity = CodeInspectionSeverity.Error;
        }

        public string Name { get { return InspectionNames.VariableNotDeclared; } }
        public CodeInspectionType InspectionType { get { return CodeInspectionType.CodeQualityIssues; } }
        public CodeInspectionSeverity Severity { get; set; }

        public IEnumerable<CodeInspectionResultBase> GetInspectionResults(VBProjectParseResult parseResult)
        {
            var issues = parseResult.IdentifierUsageInspector.UndeclaredVariableUsages();
            foreach (var issue in issues)
            {
                yield return new VariableNotDeclaredInspectionResult(Name, Severity, issue.Context, issue.QualifiedName);
            }
        }
    }
}