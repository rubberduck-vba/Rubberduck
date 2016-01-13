using System.Collections.Generic;
using System.Linq;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.UI;
using Rubberduck.VBEditor.VBEInterfaces.RubberduckCodePane;

namespace Rubberduck.Inspections
{
    public sealed class MoveFieldCloseToUsageInspection : InspectionBase
    {
        private readonly ICodePaneWrapperFactory _wrapperFactory;

        public MoveFieldCloseToUsageInspection(RubberduckParserState state)
            : base(state)
        {
            _wrapperFactory = new CodePaneWrapperFactory();
            Severity = CodeInspectionSeverity.Suggestion;
        }

        public override string Description { get { return InspectionsUI.MoveFieldCloseToUsageInspectionName; } }
        public override CodeInspectionType InspectionType { get { return CodeInspectionType.MaintainabilityAndReadabilityIssues; } }

        public override IEnumerable<CodeInspectionResultBase> GetInspectionResults()
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
                        declaration.References.Any(r => r.ParentScope != firstReference.ParentScope))
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
                        new MoveFieldCloseToUsageInspectionResult(this, issue, State, _wrapperFactory, new MessageBox()));
        }

        private Declaration ParentDeclaration(IdentifierReference reference)
        {
            return UserDeclarations.Single(d => d.Scope == reference.ParentScope && d.Project == reference.QualifiedModuleName.Project);
        }
    }
}
