using System.Collections.Generic;
using System.Linq;
using Antlr4.Runtime;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Symbols;
using Rubberduck.UI;

namespace Rubberduck.Inspections
{
    public class ObsoleteGlobalInspection : IInspection
    {
        public ObsoleteGlobalInspection()
        {
            Severity = CodeInspectionSeverity.Suggestion;
        }

        public string Name { get { return RubberduckUI.ObsoleteGlobal; } }
        public CodeInspectionType InspectionType { get { return CodeInspectionType.LanguageOpportunities; } }
        public CodeInspectionSeverity Severity { get; set; }

        public IEnumerable<CodeInspectionResultBase> GetInspectionResults(VBProjectParseResult parseResult)
        {
            var issues = from item in parseResult.Declarations.Items
                         where !item.IsBuiltIn && item.Accessibility == Accessibility.Global
                         && item.Context != null
                         select new ObsoleteGlobalInspectionResult(Name, Severity, new QualifiedContext<ParserRuleContext>(item.QualifiedName.QualifiedModuleName, item.Context));

            return issues;
        }
    }
}