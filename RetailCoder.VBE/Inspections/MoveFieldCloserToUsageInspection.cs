using System.Collections.Generic;
using System.Linq;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.UI;
using Rubberduck.VBEditor.VBEInterfaces.RubberduckCodePane;

namespace Rubberduck.Inspections
{
    public sealed class MoveFieldCloserToUsageInspection : InspectionBase
    {
        private readonly ICodePaneWrapperFactory _wrapperFactory;

        public MoveFieldCloserToUsageInspection(RubberduckParserState state)
            : base(state, CodeInspectionSeverity.Suggestion)
        {
            _wrapperFactory = new CodePaneWrapperFactory();
        }

        public override string Meta { get { return InspectionsUI.MoveFieldCloserToUsageInspectionMeta; } }
        public override string Description { get { return InspectionsUI.MoveFieldCloserToUsageInspectionName; } }
        public override CodeInspectionType InspectionType { get { return CodeInspectionType.MaintainabilityAndReadabilityIssues; } }

        public override IEnumerable<InspectionResultBase> GetInspectionResults()
        {
            return UserDeclarations
                .Where(declaration =>
                {

                    if (declaration.DeclarationType != DeclarationType.Variable ||
                        !new[] {DeclarationType.Class, DeclarationType.Module}.Contains(declaration.ParentDeclaration.DeclarationType))
                    {
                        return false;
                    }

                    var firstReference = declaration.References.FirstOrDefault();

                    if (firstReference == null ||
                        declaration.References.Any(r => r.ParentScoping != firstReference.ParentScoping))
                    {
                        return false;
                    }

                    var parentDeclaration = ParentDeclaration(firstReference);

                    return parentDeclaration != null &&
                           !new[]
                           {
                               DeclarationType.PropertyGet,
                               DeclarationType.PropertyLet,
                               DeclarationType.PropertySet
                           }.Contains(parentDeclaration.DeclarationType);
                })
                .Select(issue =>
                        new MoveFieldCloserToUsageInspectionResult(this, issue, State, _wrapperFactory, new MessageBox()));
        }

        private Declaration ParentDeclaration(IdentifierReference reference)
        {
            var declarationTypes = new[] {DeclarationType.Function, DeclarationType.Procedure, DeclarationType.Property};

            return UserDeclarations.SingleOrDefault(d =>
                        reference.ParentScoping.Equals(d) && declarationTypes.Contains(d.DeclarationType) &&
                        d.Project == reference.QualifiedModuleName.Project);
        }
    }
}
