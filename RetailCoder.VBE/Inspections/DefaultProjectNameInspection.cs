using System.Collections.Generic;
using System.Linq;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.UI;
using Rubberduck.VBEditor.VBEInterfaces.RubberduckCodePane;

namespace Rubberduck.Inspections
{
    public class DefaultProjectNameInspection : IInspection
    {
        private readonly ICodePaneWrapperFactory _wrapperFactory;

        public DefaultProjectNameInspection()
        {
            _wrapperFactory = new CodePaneWrapperFactory();
            Severity = CodeInspectionSeverity.Suggestion;
        }

        public string Name { get { return "DefaultProjectNameInspection"; } }
        public string Description { get { return RubberduckUI.GenericProjectName_; } }
        public CodeInspectionType InspectionType { get { return CodeInspectionType.MaintainabilityAndReadabilityIssues; } }
        public CodeInspectionSeverity Severity { get; set; }

        public IEnumerable<CodeInspectionResultBase> GetInspectionResults(RubberduckParserState parseResult)
        {
            var issues = parseResult.AllDeclarations
                            .Where(declaration => !declaration.IsBuiltIn 
                                                && declaration.DeclarationType == DeclarationType.Project
                                                && declaration.IdentifierName.StartsWith("VBAProject"))
                            .Select(issue => new DefaultProjectNameInspectionResult(this, issue, parseResult, _wrapperFactory))
                            .ToList();

            return issues;
        }
    }
}
