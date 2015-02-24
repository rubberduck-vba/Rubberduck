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

        public string Name { get { return InspectionNames.VariableNotDeclared_; } }
        public CodeInspectionType InspectionType { get { return CodeInspectionType.CodeQualityIssues; } }
        public CodeInspectionSeverity Severity { get; set; }

        public IEnumerable<CodeInspectionResultBase> GetInspectionResults(VBProjectParseResult parseResult)
        {
            var issues = parseResult.IdentifierUsageInspector.UndeclaredVariableUsages();
            foreach (var issue in issues)
            {
                yield return new VariableNotDeclaredInspectionResult(string.Format(Name, issue.Context.GetText()), Severity, issue.Context, issue.QualifiedName);
            }
        }
    }
}