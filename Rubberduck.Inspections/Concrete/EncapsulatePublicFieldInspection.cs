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

        protected override IEnumerable<IInspectionResult> DoGetInspectionResults()
        {
            // we're creating a public field for every control on a form, needs to be ignored.
            var fields = State.DeclarationFinder.UserDeclarations(DeclarationType.Variable)
                .Where(item => !IsIgnoringInspectionResultFor(item, AnnotationName)
                               && item.Accessibility == Accessibility.Public
                               && (item.DeclarationType != DeclarationType.Control))
                .ToList();

            return fields
                .Select(issue => new DeclarationInspectionResult(this,
                                                      string.Format(InspectionsUI.EncapsulatePublicFieldInspectionResultFormat, issue.IdentifierName),
                                                      issue))
                .ToList();
        }
    }
}
