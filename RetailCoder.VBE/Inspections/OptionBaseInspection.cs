using System.Collections.Generic;
using System.Linq;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.UI;
using Rubberduck.VBEditor.VBEInterfaces.RubberduckCodePane;

namespace Rubberduck.Inspections
{
    public class OptionBaseInspection : IInspection
    {
        private readonly ICodePaneWrapperFactory _wrapperFactory;

        public OptionBaseInspection()
        {
            _wrapperFactory = new CodePaneWrapperFactory();
            Severity = CodeInspectionSeverity.Warning;
        }

        public string Name { get { return "OptionBaseInspection"; } }
        public string Meta { get { return InspectionsUI.ResourceManager.GetString(Name + "Meta"); } }
        public string Description { get { return RubberduckUI.OptionBase; } }
        public CodeInspectionType InspectionType { get { return CodeInspectionType.MaintainabilityAndReadabilityIssues; } }
        public CodeInspectionSeverity Severity { get; set; }

        private string AnnotationName { get { return Name.Replace("Inspection", string.Empty); } }

        public IEnumerable<CodeInspectionResultBase> GetInspectionResults(RubberduckParserState state)
        {
            var options = state.AllUserDeclarations
                .Where(declaration => !declaration.IsInspectionDisabled(AnnotationName)
                                      && declaration.DeclarationType == DeclarationType.ModuleOption
                                      && declaration.Context is VBAParser.OptionBaseStmtContext)
                .ToList();

            if (!options.Any())
            {
                return new List<CodeInspectionResultBase>();
            }

            var issues = options.Where(option => ((VBAParser.OptionBaseStmtContext)option.Context).INTEGERLITERAL().GetText() == "1")
                                .Select(issue => new OptionBaseInspectionResult(this, issue.QualifiedName.QualifiedModuleName));

            return issues;
        }
    }
}