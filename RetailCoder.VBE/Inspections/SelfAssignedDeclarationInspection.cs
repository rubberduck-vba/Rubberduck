using System.Collections.Generic;
using System.Linq;
using Rubberduck.Parsing.VBA;
using Rubberduck.Parsing.Symbols;

namespace Rubberduck.Inspections
{
    public sealed class SelfAssignedDeclarationInspection : InspectionBase
    {
        public SelfAssignedDeclarationInspection(RubberduckParserState state)
            : base(state, CodeInspectionSeverity.Hint)
        {
        }

        public override string Meta { get { return InspectionsUI.SelfAssignedDeclarationInspectionMeta; } }
        public override string Description { get { return InspectionsUI.SelfAssignedDeclarationInspectionResultFormat; } }
        public override CodeInspectionType InspectionType { get { return CodeInspectionType.CodeQualityIssues; } }

        public override IEnumerable<InspectionResultBase> GetInspectionResults()
        {
            return UserDeclarations
                .Where(declaration => declaration.IsSelfAssigned 
                    && declaration.DeclarationType == DeclarationType.Variable
                    && declaration.ParentScope == declaration.QualifiedName.QualifiedModuleName.ToString()
                    && declaration.ParentDeclaration != null
                    && declaration.ParentDeclaration.DeclarationType == DeclarationType.Class)
                .Select(issue => new SelfAssignedDeclarationInspectionResult(this, issue));
        }
    }
}
