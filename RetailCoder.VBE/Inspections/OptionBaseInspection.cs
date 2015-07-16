using System.Collections.Generic;
using System.Linq;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.UI;
using Rubberduck.VBEditor.VBEInterfaces.RubberduckCodePane;

namespace Rubberduck.Inspections
{
    public class OptionBaseInspection : IInspection
    {
        private readonly IRubberduckCodePaneFactory _factory;

        public OptionBaseInspection()
        {
            _factory = new RubberduckCodePaneFactory();
            Severity = CodeInspectionSeverity.Warning;
        }

        public string Name { get { return "OptionBaseInspection"; } }
        public string Description { get { return RubberduckUI.OptionBase; } }
        public CodeInspectionType InspectionType { get { return CodeInspectionType.MaintainabilityAndReadabilityIssues; } }
        public CodeInspectionSeverity Severity { get; set; }

        public IEnumerable<CodeInspectionResultBase> GetInspectionResults(VBProjectParseResult parseResult)
        {
            var options = parseResult.Declarations.Items
                .Where(declaration => !declaration.IsBuiltIn
                                      && declaration.DeclarationType == DeclarationType.ModuleOption
                                      && declaration.Context is VBAParser.OptionBaseStmtContext)
                .ToList();

            if (!options.Any())
            {
                return new List<CodeInspectionResultBase>();
            }

            var issues = options.Where(option => ((VBAParser.OptionBaseStmtContext)option.Context).INTEGERLITERAL().GetText() == "1")
                                .Select(issue => new OptionBaseInspectionResult(Description, Severity, issue.QualifiedName.QualifiedModuleName, _factory));

            return issues;
        }
    }
}