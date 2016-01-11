using System.Collections.Generic;
using System.Linq;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor.VBEInterfaces.RubberduckCodePane;

namespace Rubberduck.Inspections
{
    public class EncapsulatePublicFieldInspection : IInspection
    {
        private readonly ICodePaneWrapperFactory _wrapperFactory;

        public EncapsulatePublicFieldInspection()
        {
            _wrapperFactory = new CodePaneWrapperFactory();
            Severity = CodeInspectionSeverity.Suggestion;
        }

        public string Name { get { return "EncapsulatePublicFieldInspection"; } }
        public string Meta { get { return InspectionsUI.ResourceManager.GetString(Name + "Meta"); } }
        public string Description { get { return InspectionsUI.EncapsulatePublicFieldInspectionName; } }
        public CodeInspectionType InspectionType { get { return CodeInspectionType.MaintainabilityAndReadabilityIssues; } }
        public CodeInspectionSeverity Severity { get; set; }

        public IEnumerable<CodeInspectionResultBase> GetInspectionResults(RubberduckParserState state)
        {
            var issues = state.AllDeclarations
                            .Where(declaration => !declaration.IsBuiltIn 
                                                && declaration.DeclarationType == DeclarationType.Variable
                                                && declaration.Accessibility == Accessibility.Public)
                            .Select(issue => new EncapsulatePublicFieldInspectionResult(this, issue, state, _wrapperFactory))
                            .ToList();

            return issues;
        }
    }
}
