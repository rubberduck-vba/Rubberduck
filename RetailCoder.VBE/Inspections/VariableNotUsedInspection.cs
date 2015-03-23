using System.Collections.Generic;
using System.Linq;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Symbols;

namespace Rubberduck.Inspections
{
    public class VariableNotUsedInspection : IInspection
    {
        public VariableNotUsedInspection()
        {
            Severity = CodeInspectionSeverity.Hint;
        }

        public string Name { get { return InspectionNames.VariableNotUsed_; } }
        public CodeInspectionType InspectionType { get { return CodeInspectionType.CodeQualityIssues; } }
        public CodeInspectionSeverity Severity { get; set; }

        public IEnumerable<CodeInspectionResultBase> GetInspectionResults(VBProjectParseResult parseResult)
        {
            var declarations = parseResult.Declarations.Items.Where(declaration =>
                (declaration.DeclarationType == DeclarationType.Variable)
                && !declaration.References.Any(reference => !reference.IsAssignment));

            foreach (var issue in declarations)
            {
                yield return new VariableNotUsedInspectionResult(string.Format(Name, issue.IdentifierName), Severity, issue.Context, issue.QualifiedName.ModuleScope);
            }
        }
    }
}