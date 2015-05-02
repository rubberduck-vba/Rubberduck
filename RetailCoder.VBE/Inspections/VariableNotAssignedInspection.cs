using System.Collections.Generic;
using System.Linq;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Symbols;

namespace Rubberduck.Inspections
{
    public class VariableNotAssignedInspection : IInspection
    {
        public VariableNotAssignedInspection()
        {
            Severity = CodeInspectionSeverity.Warning;
        }

        public string Name { get { return InspectionNames.VariableNotAssigned_; } }
        public CodeInspectionType InspectionType { get { return CodeInspectionType.CodeQualityIssues; } }
        public CodeInspectionSeverity Severity { get; set; }

        public IEnumerable<CodeInspectionResultBase> GetInspectionResults(VBProjectParseResult parseResult)
        {
            var declarations = parseResult.Declarations.Items.Where(declaration =>
                !declaration.IsBuiltIn 
                && declaration.DeclarationType == DeclarationType.Variable
                && !declaration.IsArray() // ignore arrays... not ideal though
                && !parseResult.Declarations.Items.Any(item => 
                    item.IdentifierName == declaration.AsTypeName 
                    && item.DeclarationType == DeclarationType.UserDefinedType) // UDT variables don't need to be assigned
                && !declaration.IsSelfAssigned
                && !declaration.References.Any(reference => reference.IsAssignment));

            foreach (var issue in declarations)
            {
                yield return new IdentifierNotAssignedInspectionResult(string.Format(Name, issue.IdentifierName), Severity, issue.Context, issue.QualifiedName.QualifiedModuleName);
            }
        }
    }
}