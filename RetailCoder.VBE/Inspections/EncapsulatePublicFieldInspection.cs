using System.Collections.Generic;
using System.Linq;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor.VBEInterfaces.RubberduckCodePane;

namespace Rubberduck.Inspections
{
    public sealed class EncapsulatePublicFieldInspection : InspectionBase
    {
        private readonly ICodePaneWrapperFactory _wrapperFactory;

        public EncapsulatePublicFieldInspection(RubberduckParserState state)
            : base(state, CodeInspectionSeverity.Suggestion)
        {
            _wrapperFactory = new CodePaneWrapperFactory();
        }

        public override string Meta { get { return InspectionsUI.EncapsulatePublicFieldInspectionMeta; } }
        public override string Description { get { return InspectionsUI.EncapsulatePublicFieldInspectionName; } }
        public override CodeInspectionType InspectionType { get { return CodeInspectionType.MaintainabilityAndReadabilityIssues; } }

        public override IEnumerable<InspectionResultBase> GetInspectionResults()
        {
            var issues = UserDeclarations
                            .Where(declaration => declaration.DeclarationType == DeclarationType.Variable
                                                && declaration.Accessibility == Accessibility.Public)
                            .Select(issue => new EncapsulatePublicFieldInspectionResult(this, issue, State, _wrapperFactory))
                            .ToList();

            return issues;
        }
    }
}
