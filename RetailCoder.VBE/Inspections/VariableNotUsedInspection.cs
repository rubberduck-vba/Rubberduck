using System.Collections.Generic;
using System.Linq;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.Inspections
{
    public sealed class VariableNotUsedInspection : InspectionBase
    {
        public VariableNotUsedInspection(RubberduckParserState state)
            : base(state)
        {
        }

        public override string Meta { get { return InspectionsUI.VariableNotUsedInspectionMeta; } }
        public override string Description { get { return InspectionsUI.VariableNotUsedInspectionName; } }
        public override CodeInspectionType InspectionType { get { return CodeInspectionType.CodeQualityIssues; } }

        public override IEnumerable<InspectionResultBase> GetInspectionResults()
        {
            var declarations = UserDeclarations.Where(declaration =>
                !declaration.IsWithEvents
                && declaration.DeclarationType == DeclarationType.Variable
                && declaration.References.All(reference => reference.IsAssignment));

            return declarations.Select(issue => 
                new IdentifierNotUsedInspectionResult(this, issue, ((dynamic)issue.Context).identifier(), issue.QualifiedName.QualifiedModuleName));
        }
    }
}