using System.Collections.Generic;
using System.Linq;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.Inspections
{
    public sealed class VariableNotAssignedInspection : InspectionBase
    {
        public VariableNotAssignedInspection(RubberduckParserState state)
            : base(state)
        {
        }

        public override string Meta { get { return InspectionsUI.VariableNotAssignedInspectionMeta; } }
        public override string Description { get { return InspectionsUI.VariableNotAssignedInspectionName; } }
        public override CodeInspectionType InspectionType { get { return CodeInspectionType.CodeQualityIssues; } }

        public override IEnumerable<InspectionResultBase> GetInspectionResults()
        {
            var items = UserDeclarations.ToList();

            // ignore arrays. todo: ArrayIndicesNotAccessedInspection
            var arrays = items.Where(declaration =>
                declaration.DeclarationType == DeclarationType.Variable
                && declaration.IsArray()).ToList();

            var declarations = items.Except(arrays).Where(declaration =>
                declaration.DeclarationType == DeclarationType.Variable
                && !declaration.IsWithEvents
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
