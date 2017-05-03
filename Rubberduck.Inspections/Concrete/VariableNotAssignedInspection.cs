using System.Collections.Generic;
using System.Linq;
using Rubberduck.Common;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Inspections.Resources;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.Inspections.Concrete
{
    public sealed class VariableNotAssignedInspection : InspectionBase
    {
        public VariableNotAssignedInspection(RubberduckParserState state)
            : base(state) { }

        public override CodeInspectionType InspectionType => CodeInspectionType.CodeQualityIssues;

        public override IEnumerable<IInspectionResult> GetInspectionResults()
        {
            var items = UserDeclarations.ToList();

            // ignore arrays. todo: ArrayIndicesNotAccessedInspection
            var arrays = items.Where(declaration =>
                declaration.DeclarationType == DeclarationType.Variable
                && declaration.IsArray).ToList();

            var declarations = items.Except(arrays).Where(declaration =>
                declaration.DeclarationType == DeclarationType.Variable
                && !declaration.IsWithEvents
                && !items.Any(item => 
                    item.IdentifierName == declaration.AsTypeName 
                    && item.DeclarationType == DeclarationType.UserDefinedType) // UDT variables don't need to be assigned
                && !declaration.IsSelfAssigned
                && !declaration.References.Any(reference => reference.IsAssignment));

            return declarations.Select(issue => 
                new InspectionResult(this, string.Format(InspectionsUI.VariableNotAssignedInspectionResultFormat, issue.IdentifierName).Capitalize(), issue));
        }
    }
}
