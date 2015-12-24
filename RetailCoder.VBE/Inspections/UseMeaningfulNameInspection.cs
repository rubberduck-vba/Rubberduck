using System.Collections.Generic;
using System.Linq;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor.VBEInterfaces.RubberduckCodePane;

namespace Rubberduck.Inspections
{
    public class UseMeaningfulNameInspection : IInspection
    {
        private readonly ICodePaneWrapperFactory _wrapperFactory;

        public UseMeaningfulNameInspection()
        {
            _wrapperFactory = new CodePaneWrapperFactory();
            Severity = CodeInspectionSeverity.Suggestion;
        }

        public string Name { get { return "UseMeaningfulNameInspection"; } }
        public string Description { get { return InspectionsUI.UseMeaningfulNameInspection; } }
        public CodeInspectionType InspectionType { get { return CodeInspectionType.MaintainabilityAndReadabilityIssues; } }
        public CodeInspectionSeverity Severity { get; set; }

        public IEnumerable<CodeInspectionResultBase> GetInspectionResults(RubberduckParserState state)
        {
            var issues = state.AllDeclarations
                            .Where(declaration => !declaration.IsBuiltIn &&
                                                  (declaration.IdentifierName.Length < 3 ||
                                                   char.IsDigit(declaration.IdentifierName.Last()) ||
                                                   !declaration.IdentifierName.Any(c => new[] {'a', 'e', 'i', 'o', 'u', 'y'}.Contains(c))
                                                  ))
                            .Select(issue => new UseMeaningfulNameInspectionResult(this, issue, state, _wrapperFactory))
                            .ToList();

            return issues;
        }
    }
}