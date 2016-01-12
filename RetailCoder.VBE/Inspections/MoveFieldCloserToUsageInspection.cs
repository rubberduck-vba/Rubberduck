using System.Collections.Generic;
using System.Linq;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.UI;
using Rubberduck.VBEditor.VBEInterfaces.RubberduckCodePane;

namespace Rubberduck.Inspections
{
    public class MoveFieldCloseToUsageInspection : IInspection
    {
        private readonly ICodePaneWrapperFactory _wrapperFactory;

        public MoveFieldCloseToUsageInspection()
        {
            _wrapperFactory = new CodePaneWrapperFactory();
            Severity = CodeInspectionSeverity.Suggestion;
        }

        public string Name { get { return "MoveFieldCloseToUsageInspection"; } }
        public string Meta { get { return InspectionsUI.ResourceManager.GetString(Name + "Meta"); } }
        public string Description { get { return InspectionsUI.MoveFieldCloseToUsageInspectionName; } }
        public CodeInspectionType InspectionType { get { return CodeInspectionType.MaintainabilityAndReadabilityIssues; } }
        public CodeInspectionSeverity Severity { get; set; }

        public IEnumerable<CodeInspectionResultBase> GetInspectionResults(RubberduckParserState state)
        {
            return state.AllDeclarations
                .Where(declaration =>
                {

                    if (declaration.IsBuiltIn ||
                        declaration.DeclarationType != DeclarationType.Variable)
                    {
                        return false;
                    }

                    var firstReference = declaration.References.FirstOrDefault();

                    return firstReference != null &&
                           declaration.References.All(r => r.ParentScope == firstReference.ParentScope);
                })
                .Select(issue =>
                        new MoveFieldCloseToUsageInspectionResult(this, issue, state, _wrapperFactory, new MessageBox()));
        }
    }
}
