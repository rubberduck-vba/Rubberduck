using System.Collections.Generic;
using Rubberduck.VBA.Nodes;

namespace Rubberduck.Inspections
{
    public class ParameterNotUsedInspection : IInspection
    {
        public ParameterNotUsedInspection()
        {
            Severity = CodeInspectionSeverity.Warning;
        }

        public string Name { get { return InspectionNames.ParameterNotUsed; } }
        public CodeInspectionType InspectionType { get { return CodeInspectionType.CodeQualityIssues; } }
        public CodeInspectionSeverity Severity { get; set; }

        public IEnumerable<CodeInspectionResultBase> GetInspectionResults(IEnumerable<VBComponentParseResult> parseResult)
        {
            var inspector = new IdentifierUsageInspector(parseResult);
            var issues = inspector.UnusedParameters();

            foreach (var issue in issues)
            {
                yield return new ParameterNotUsedInspectionResult(Name, Severity, issue.Context, issue.MemberName);
            }
        }
    }
}