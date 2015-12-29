using System.Collections.Generic;
using System.Linq;
using Antlr4.Runtime;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.UI;

namespace Rubberduck.Inspections
{
    public class ObsoleteGlobalInspection : IInspection
    {
        public ObsoleteGlobalInspection()
        {
            Severity = CodeInspectionSeverity.Suggestion;
        }

        public string Name { get { return "ObsoleteGlobalInspection"; } }
        public string Meta { get { return InspectionsUI.ResourceManager.GetString(Name + "Meta"); } }
        public string Description { get { return RubberduckUI.ObsoleteGlobal; } }
        public CodeInspectionType InspectionType { get { return CodeInspectionType.LanguageOpportunities; } }
        public CodeInspectionSeverity Severity { get; set; }

        public IEnumerable<CodeInspectionResultBase> GetInspectionResults(RubberduckParserState state)
        {
            var issues = from item in state.AllUserDeclarations
                         where item.Accessibility == Accessibility.Global && item.Context != null
                         select new ObsoleteGlobalInspectionResult(this, Description, new QualifiedContext<ParserRuleContext>(item.QualifiedName.QualifiedModuleName, item.Context));

            return issues;
        }
    }
}