using System.Collections.Generic;
using System.Linq;
using Antlr4.Runtime;
using Rubberduck.Parsing;
using Rubberduck.Parsing.VBA;
using Rubberduck.UI;

namespace Rubberduck.Inspections
{
    public class ObsoleteLetStatementInspection : IInspection
    {
        public ObsoleteLetStatementInspection()
        {
            Severity = CodeInspectionSeverity.Suggestion;
        }

        public string Name { get { return "ObsoleteLetStatementInspection"; } }
        public string Description { get { return RubberduckUI.ObsoleteLet; } }
        public CodeInspectionType InspectionType { get { return CodeInspectionType.LanguageOpportunities; } }
        public CodeInspectionSeverity Severity { get; set; }

        public IEnumerable<CodeInspectionResultBase> GetInspectionResults(RubberduckParserState state)
        {
            return state.ObsoleteLetContexts.Select(context =>
                new ObsoleteLetStatementUsageInspectionResult(this, new QualifiedContext<ParserRuleContext>(context.ModuleName, context.Context)));
        }
    }
}