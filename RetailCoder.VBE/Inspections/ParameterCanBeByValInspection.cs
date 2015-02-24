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

        public string Name { get { return InspectionNames.ParameterCanBeByVal_; } }
        public CodeInspectionType InspectionType { get { return CodeInspectionType.CodeQualityIssues; } }
        public CodeInspectionSeverity Severity { get; set; }

        public IEnumerable<CodeInspectionResultBase> GetInspectionResults(VBProjectParseResult parseResult)
        {
            var issues = parseResult.IdentifierUsageInspector.UnassignedByRefParameters();

            foreach (var issue in issues)
            {
                yield return new ParameterCanBeByValInspectionResult(string.Format(Name, issue.Context.GetText()), Severity, issue.Context, issue.MemberName);
            }
        }
    }
}