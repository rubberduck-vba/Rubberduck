using System.Collections.Generic;
using System.Linq;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Results;
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
            // ignore arrays. todo: ArrayIndicesNotAccessedInspection
            var arrays = State.DeclarationFinder.UserDeclarations(DeclarationType.Variable).Where(declaration => declaration.IsArray);

            var declarations = State.DeclarationFinder.UserDeclarations(DeclarationType.Variable)
                .Except(arrays)
                .Where(declaration =>
                    !declaration.IsWithEvents
                    && State.DeclarationFinder.MatchName(declaration.AsTypeName).All(item => item.DeclarationType != DeclarationType.UserDefinedType) // UDT variables don't need to be assigned
                    && !declaration.IsSelfAssigned
                    && !declaration.References.Any(reference => reference.IsAssignment))
                .Where(result => !IsIgnoringInspectionResultFor(result, AnnotationName));

            return declarations.Select(issue => 
                new DeclarationInspectionResult(this, string.Format(InspectionsUI.VariableNotAssignedInspectionResultFormat, issue.IdentifierName), issue));
        }
    }
}
