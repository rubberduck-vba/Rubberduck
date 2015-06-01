using System.Collections.Generic;
using System.Linq;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Symbols;
using Rubberduck.UI;

namespace Rubberduck.Inspections
{
    public class VariableNotAssignedInspection : IInspection
    {
        public VariableNotAssignedInspection()
        {
            Severity = CodeInspectionSeverity.Warning;
        }

        public string Name { get { return "VariableNotAssignedInspection"; } }
        public string Description { get { return RubberduckUI.VariableNotAssigned_; } }
        public CodeInspectionType InspectionType { get { return CodeInspectionType.CodeQualityIssues; } }
        public CodeInspectionSeverity Severity { get; set; }

        public IEnumerable<CodeInspectionResultBase> GetInspectionResults(VBProjectParseResult parseResult)
        {
            // ignore arrays. todo: ArrayIndicesNotAccessedInspection
            var arrays = parseResult.Declarations.Items.Where(declaration =>
                declaration.DeclarationType == DeclarationType.Variable
                && declaration.IsArray()).ToList();

            var declarations = parseResult.Declarations.Items.Where(declaration =>
                declaration.DeclarationType == DeclarationType.Variable
                && !declaration.IsBuiltIn 
                && !arrays.Contains(declaration)
                && !parseResult.Declarations.Items.Any(item => 
                    item.IdentifierName == declaration.AsTypeName 
                    && item.DeclarationType == DeclarationType.UserDefinedType) // UDT variables don't need to be assigned
                && !declaration.IsSelfAssigned
                && !declaration.References.Any(reference => reference.IsAssignment));

            foreach (var issue in declarations)
            {
                yield return new IdentifierNotAssignedInspectionResult(string.Format(Description, issue.IdentifierName), Severity, ((dynamic)issue.Context).ambiguousIdentifier(), issue.QualifiedName.QualifiedModuleName);
            }
        }
    }
}