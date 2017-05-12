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
    public sealed class EncapsulatePublicFieldInspection : InspectionBase
    {
        public EncapsulatePublicFieldInspection(RubberduckParserState state)
            : base(state, CodeInspectionSeverity.Suggestion) { }

        public override CodeInspectionType InspectionType => CodeInspectionType.MaintainabilityAndReadabilityIssues;

        public override IEnumerable<IInspectionResult> GetInspectionResults()
        {
            // we're creating a public field for every control on a form, needs to be ignored.
            var msForms = State.DeclarationFinder.FindProject("MSForms");
            Declaration control = null;
            if (msForms != null)
            {
                control = State.DeclarationFinder.FindClassModule("Control", msForms, true);
            }

            var fields = State.DeclarationFinder.UserDeclarations(DeclarationType.Variable)
                .Where(item => !IsIgnoringInspectionResultFor(item, AnnotationName)
                               && item.Accessibility == Accessibility.Public
                               && (control == null || !Equals(item.AsTypeDeclaration, control)))
                .ToList();

            return fields
                .Select(issue => new DeclarationInspectionResult(this,
                                                      string.Format(InspectionsUI.EncapsulatePublicFieldInspectionResultFormat, issue.IdentifierName),
                                                      issue))
                .ToList();
        }
    }
}
