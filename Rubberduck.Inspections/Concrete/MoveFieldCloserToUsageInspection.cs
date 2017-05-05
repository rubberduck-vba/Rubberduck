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
    public sealed class MoveFieldCloserToUsageInspection : InspectionBase
    {
        public MoveFieldCloserToUsageInspection(RubberduckParserState state)
            : base(state, CodeInspectionSeverity.Suggestion) { }

        public override CodeInspectionType InspectionType => CodeInspectionType.MaintainabilityAndReadabilityIssues;

        public override IEnumerable<IInspectionResult> GetInspectionResults()
        {
            return State.DeclarationFinder.UserDeclarations(DeclarationType.Variable)
                .Where(declaration =>
                {
                    if (declaration.IsWithEvents
                        || !new[] {DeclarationType.ClassModule, DeclarationType.ProceduralModule}.Contains(declaration.ParentDeclaration.DeclarationType)
                        || IsIgnoringInspectionResultFor(declaration, AnnotationName))
                    {
                        return false;
                    }

                    var asType = declaration.AsTypeDeclaration;
                    if (asType != null && asType.ProjectName.Equals("Rubberduck") &&
                        (asType.IdentifierName.Equals("PermissiveAssertClass") || asType.IdentifierName.Equals("AssertClass")))
                    {
                        return false;
                    }

                    var firstReference = declaration.References.FirstOrDefault();

                    if (firstReference == null ||
                        declaration.References.Any(r => r.ParentScoping != firstReference.ParentScoping))
                    {
                        return false;
                    }

                    var parentDeclaration = ParentDeclaration(firstReference);

                    return parentDeclaration != null &&
                           !new[]
                           {
                               DeclarationType.PropertyGet,
                               DeclarationType.PropertyLet,
                               DeclarationType.PropertySet
                           }.Contains(parentDeclaration.DeclarationType);
                })
                .Select(issue =>
                        new DeclarationInspectionResult(this, string.Format(InspectionsUI.MoveFieldCloserToUsageInspectionResultFormat, issue.IdentifierName), issue));
        }

        private Declaration ParentDeclaration(IdentifierReference reference)
        {
            var declarationTypes = new[] {DeclarationType.Function, DeclarationType.Procedure, DeclarationType.Property};

            return UserDeclarations.SingleOrDefault(d =>
                        reference.ParentScoping.Equals(d) && declarationTypes.Contains(d.DeclarationType) &&
                        d.QualifiedName.QualifiedModuleName.Equals(reference.QualifiedModuleName));
        }
    }
}
