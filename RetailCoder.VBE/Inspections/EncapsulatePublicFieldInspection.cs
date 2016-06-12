using System.Collections.Generic;
using System.Linq;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.Inspections
{
    public sealed class EncapsulatePublicFieldInspection : InspectionBase
    {
        public EncapsulatePublicFieldInspection(RubberduckParserState state)
            : base(state, CodeInspectionSeverity.Suggestion)
        {
        }

        public override string Meta { get { return InspectionsUI.EncapsulatePublicFieldInspectionMeta; } }
        public override string Description { get { return InspectionsUI.EncapsulatePublicFieldInspectionName; } }
        public override CodeInspectionType InspectionType { get { return CodeInspectionType.MaintainabilityAndReadabilityIssues; } }

        public override IEnumerable<InspectionResultBase> GetInspectionResults()
        {
            var issues = UserDeclarations
                            .Where(declaration => declaration.DeclarationType == DeclarationType.Variable
                                                && declaration.Accessibility == Accessibility.Public)
                            .Select(issue => new EncapsulatePublicFieldInspectionResult(this, issue, State))
                            .ToList();

            return issues;
        }
    }
}
