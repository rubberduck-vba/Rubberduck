using System.Collections.Generic;
using System.Linq;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Results;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.Inspections.Concrete
{
    public sealed class EncapsulatePublicFieldInspection : InspectionBase
    {
        public EncapsulatePublicFieldInspection(RubberduckParserState state)
            : base(state) { }

        protected override IEnumerable<IInspectionResult> DoGetInspectionResults()
        {
            // we're creating a public field for every control on a form, needs to be ignored.
            var fields = State.DeclarationFinder.UserDeclarations(DeclarationType.Variable)
                .Where(item => !IsIgnoringInspectionResultFor(item, AnnotationName)
                               && item.Accessibility == Accessibility.Public
                               && (item.DeclarationType != DeclarationType.Control));
            return fields
                .Select(issue => new DeclarationInspectionResult(this, string.Format(Resources.Inspections.InspectionResults.EncapsulatePublicFieldInspection, issue.IdentifierName), issue))
                .ToList();
        }
    }
}
