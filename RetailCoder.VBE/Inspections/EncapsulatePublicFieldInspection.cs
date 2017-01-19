using System.Collections.Generic;
using System.Linq;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Resources;
using Rubberduck.Inspections.Results;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using IInspectionResult = Rubberduck.Parsing.Symbols.IInspectionResult;

namespace Rubberduck.Inspections
{
    using SmartIndenter;

    public sealed class EncapsulatePublicFieldInspection : InspectionBase
    {
        private readonly IIndenter _indenter;

        public EncapsulatePublicFieldInspection(RubberduckParserState state, IIndenter indenter)
            : base(state, CodeInspectionSeverity.Suggestion)
        {
            _indenter = indenter;
        }

        public override string Meta { get { return InspectionsUI.EncapsulatePublicFieldInspectionMeta; } }
        public override string Description { get { return InspectionsUI.EncapsulatePublicFieldInspectionName; } }
        public override CodeInspectionType InspectionType { get { return CodeInspectionType.MaintainabilityAndReadabilityIssues; } }

        public override IEnumerable<IInspectionResult> GetInspectionResults()
        {
            var fields = State.DeclarationFinder.UserDeclarations(DeclarationType.Variable)
                .Where(item => !IsIgnoringInspectionResultFor(item, AnnotationName)
                               && item.Accessibility == Accessibility.Public)
                .ToList();

            return fields
                .Select(issue => new EncapsulatePublicFieldInspectionResult(this, issue, State, _indenter))
                .ToList();
        }
    }
}
