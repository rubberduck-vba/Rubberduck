using System.Collections.Generic;
using System.Linq;
using Rubberduck.Parsing.VBA;
using Rubberduck.UI;
using Rubberduck.VBEditor.VBEInterfaces.RubberduckCodePane;

namespace Rubberduck.Inspections
{
    public class UseMeaningfulNameInspection : IInspection
    {
        private readonly IMessageBox _messageBox;
        private readonly ICodePaneWrapperFactory _wrapperFactory;

        public UseMeaningfulNameInspection(IMessageBox messageBox)
        {
            _messageBox = messageBox;
            _wrapperFactory = new CodePaneWrapperFactory();
            Severity = CodeInspectionSeverity.Suggestion;
        }

        public string Name { get { return "UseMeaningfulNameInspection"; } }
        public string Meta { get { return InspectionsUI.ResourceManager.GetString(Name + "Meta"); } }
        public string Description { get { return InspectionsUI.UseMeaningfulNameInspection; } }
        public CodeInspectionType InspectionType { get { return CodeInspectionType.MaintainabilityAndReadabilityIssues; } }
        public CodeInspectionSeverity Severity { get; set; }

        public IEnumerable<CodeInspectionResultBase> GetInspectionResults(RubberduckParserState state)
        {
            var issues = state.AllUserDeclarations
                            .Where(declaration => declaration.IdentifierName.Length < 3 ||
                                                  char.IsDigit(declaration.IdentifierName.Last()) ||
                                                  !declaration.IdentifierName.Any(c => new[] {'a', 'e', 'i', 'o', 'u', 'y'}.Contains(c)))
                            .Select(issue => new UseMeaningfulNameInspectionResult(this, issue, state, _wrapperFactory, _messageBox))
                            .ToList();

            return issues;
        }
    }
}