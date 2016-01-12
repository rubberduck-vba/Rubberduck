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

                    return firstReference != null &&
                           declaration.References.All(r => r.ParentScope == firstReference.ParentScope);
                })
                .Select(issue =>
                        new MoveFieldCloseToUsageInspectionResult(this, issue, State, _wrapperFactory, new MessageBox()));
        }
    }
}
