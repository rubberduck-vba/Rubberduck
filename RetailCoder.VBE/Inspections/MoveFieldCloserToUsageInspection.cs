using System.Collections.Generic;
using System.Linq;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.UI;
using Rubberduck.VBEditor;
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
                           declaration.References.All(r => r.ParentScope == firstReference.ParentScope) &&
                           !new[]
                           {
                               DeclarationType.PropertyGet,
                               DeclarationType.PropertyLet,
                               DeclarationType.PropertySet
                           }.Contains(ParentDeclaration(firstReference).DeclarationType);
                })
                .Select(issue =>
                        new MoveFieldCloseToUsageInspectionResult(this, issue, State, _wrapperFactory, new MessageBox()));
        }

        private Declaration ParentDeclaration(IdentifierReference reference)
        {
            Declaration activeDeclaration = null;

            var activeSelection = new Selection(0, 0, int.MaxValue, int.MaxValue);

            foreach (var declaration in UserDeclarations.Where(d => d.Scope == reference.ParentScope))
            {
                if (new Selection(declaration.Context.Start.Line,
                                  declaration.Context.Start.Column,
                                  declaration.Context.Stop.Line,
                                  declaration.Context.Stop.Column)
                    .Contains(reference.Selection) &&
                    activeSelection.Contains(declaration.Selection))
                {
                    activeDeclaration = declaration;
                    activeSelection = declaration.Selection;
                }
            }

            return activeDeclaration;
        }
    }
}
