using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using Rubberduck.VBA;
using Rubberduck.VBA.Grammar;
using Rubberduck.VBA.Nodes;

namespace Rubberduck.Inspections
{
    public class VariableNotAssignedInspection : IInspection
    {
        public VariableNotAssignedInspection()
        {
            Severity = CodeInspectionSeverity.Hint;
        }

        public string Name { get { return InspectionNames.VariableNotAssigned; } }
        public CodeInspectionType InspectionType { get { return CodeInspectionType.CodeQualityIssues; } }
        public CodeInspectionSeverity Severity { get; set; }

        public IEnumerable<CodeInspectionResultBase> GetInspectionResults(IEnumerable<VBComponentParseResult> parseResult)
        {
            var inspector = new IdentifierUsageInspector(parseResult);
            var issues = inspector.UnassignedGlobals()
                  .Union(inspector.UnassignedFields())
                  .Union(inspector.UnassignedLocals());

            foreach (var issue in issues)
            {
                yield return new VariableNotAssignedInspectionResult(Name, Severity, issue.Context, issue.QualifiedName);
            }
        }
    }
}