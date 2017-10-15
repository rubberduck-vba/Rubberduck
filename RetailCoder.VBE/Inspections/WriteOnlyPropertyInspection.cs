using System.Collections.Generic;
using System.Linq;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Resources;
using Rubberduck.Inspections.Results;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.Inspections
{
    public sealed class WriteOnlyPropertyInspection : InspectionBase
    {
        public WriteOnlyPropertyInspection(RubberduckParserState state)
            : base(state, CodeInspectionSeverity.Suggestion)
        {
        }

        public override string Meta { get { return InspectionsUI.WriteOnlyPropertyInspectionMeta; } }
        public override string Description { get { return InspectionsUI.WriteOnlyPropertyInspectionName; } }
        public override CodeInspectionType InspectionType { get { return CodeInspectionType.CodeQualityIssues; } }

        public override IEnumerable<InspectionResultBase> GetInspectionResults()
        {
            var declarations = UserDeclarations.ToList();
            var setters = declarations
                .Where(item => 
                       (item.Accessibility == Accessibility.Implicit || 
                        item.Accessibility == Accessibility.Public || 
                        item.Accessibility == Accessibility.Global)
                    && (item.DeclarationType == DeclarationType.PropertyLet ||
                        item.DeclarationType == DeclarationType.PropertySet)
                    && !declarations.Where(declaration => declaration.IdentifierName == item.IdentifierName)
                        .Any(accessor => !accessor.IsBuiltIn && accessor.DeclarationType == DeclarationType.PropertyGet))
                .GroupBy(item => new {item.QualifiedName, item.DeclarationType})
                .Select(grouping => grouping.First()); // don't get both Let and Set accessors

            return setters.Select(setter =>
                new WriteOnlyPropertyInspectionResult(this, setter));
        }
    }
}
