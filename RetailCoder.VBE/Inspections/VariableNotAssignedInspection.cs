using System.Collections.Generic;
using System.Linq;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
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

        public IEnumerable<CodeInspectionResultBase> GetInspectionResults(RubberduckParserState state)
        {
            var items = state.AllDeclarations.ToList();

            // ignore arrays. todo: ArrayIndicesNotAccessedInspection
            var arrays = items.Where(declaration =>
                declaration.DeclarationType == DeclarationType.Variable
                && declaration.IsArray()).ToList();

            var declarations = items.Where(declaration =>
                declaration.DeclarationType == DeclarationType.Variable
                && !declaration.IsBuiltIn 
                && !declaration.IsWithEvents
                && !arrays.Contains(declaration)
                && !items.Any(item => 
                    item.IdentifierName == declaration.AsTypeName 
                    && item.DeclarationType == DeclarationType.UserDefinedType) // UDT variables don't need to be assigned
                && !declaration.IsSelfAssigned
                && !declaration.References.Any(reference => reference.IsAssignment));

            return declarations.Select(issue => 
                new IdentifierNotAssignedInspectionResult(this, issue, issue.Context, issue.QualifiedName.QualifiedModuleName));
        }
    }
}