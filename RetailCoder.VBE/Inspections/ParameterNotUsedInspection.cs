using System.Collections.Generic;
using Rubberduck.Parsing;

namespace Rubberduck.Inspections
{
    public class ParameterNotUsedInspection : IInspection
    {
        public ParameterNotUsedInspection()
        {
            Severity = CodeInspectionSeverity.Hint;
        }

        public string Name { get { return InspectionNames.ParameterNotUsed_; } }
        public CodeInspectionType InspectionType { get { return CodeInspectionType.CodeQualityIssues; } }
        public CodeInspectionSeverity Severity { get; set; }

        public IEnumerable<CodeInspectionResultBase> GetInspectionResults(VBProjectParseResult parseResult)
        {
            var issues = parseResult.IdentifierUsageInspector.UnusedParameters();
            foreach (var issue in issues)
            {
                yield return new ParameterNotUsedInspectionResult(string.Format(Name, issue.Context.GetText()), Severity, issue.Context, issue.MemberName);
            }
        }
    }
}