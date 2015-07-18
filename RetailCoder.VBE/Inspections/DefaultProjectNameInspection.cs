using System.Collections.Generic;
using System.Linq;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Symbols;
using Rubberduck.UI;
using Rubberduck.VBEditor.VBEInterfaces.RubberduckCodePane;

namespace Rubberduck.Inspections
{
    public class DefaultProjectNameInspection : IInspection
    {
        private readonly IRubberduckCodePaneFactory _factory;

        public DefaultProjectNameInspection()
        {
            _factory = new RubberduckCodePaneFactory();
            Severity = CodeInspectionSeverity.Suggestion;
        }

        public string Name { get { return "DefaultProjectNameInspection"; } }
        public string Description { get { return RubberduckUI.GenericProjectName_; } }
        public CodeInspectionType InspectionType { get { return CodeInspectionType.MaintainabilityAndReadabilityIssues; } }
        public CodeInspectionSeverity Severity { get; set; }

        public IEnumerable<CodeInspectionResultBase> GetInspectionResults(VBProjectParseResult parseResult)
        {
            var issues = parseResult.Declarations.Items
                            .Where(declaration => !declaration.IsBuiltIn 
                                                && declaration.DeclarationType == DeclarationType.Project
                                                && declaration.IdentifierName.StartsWith("VBAProject"))
                            .Select(issue => new DefaultProjectNameInspectionResult(string.Format(Description, issue.IdentifierName), Severity, issue, parseResult, _factory))
                            .ToList();

            return issues;
        }
    }
}
