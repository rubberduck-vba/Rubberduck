using System.Collections.Generic;
using System.Linq;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.Inspections
{
    public sealed class ConstantNotUsedInspection : InspectionBase
    {
        public ConstantNotUsedInspection(RubberduckParserState state)
            : base(state)
        {
        }

        public override string Meta { get { return InspectionsUI.ConstantNotUsedInspectionMeta; } }
        public override string Description { get { return InspectionsUI.ConstantNotUsedInspectionName; } }
        public override CodeInspectionType InspectionType { get { return CodeInspectionType.CodeQualityIssues; } }

        public override IEnumerable<InspectionResultBase> GetInspectionResults()
        {
            var results = UserDeclarations.Where(declaration =>
                    declaration.DeclarationType == DeclarationType.Constant && !declaration.References.Any());

            return results.Select(issue => 
                new IdentifierNotUsedInspectionResult(this, issue, ((dynamic)issue.Context).identifier(), issue.QualifiedName.QualifiedModuleName)).Cast<InspectionResultBase>();
        }
    }
}