using System.Collections.Generic;
using System.Linq;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Resources;
using Rubberduck.Inspections.Results;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;

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

        public override IEnumerable<InspectionResultBase> GetInspectionResults()
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
                .Select(issue => new EncapsulatePublicFieldInspectionResult(this, issue, State, _indenter))
                .ToList();
        }
    }
}
