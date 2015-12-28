using System.Collections.Generic;
using System.Linq;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.VBA;
using Rubberduck.UI;

namespace Rubberduck.Inspections
{
    public class ObsoleteCallStatementInspection : IInspection
    {
        public ObsoleteCallStatementInspection()
        {
            Severity = CodeInspectionSeverity.Suggestion;
        }

        public string Name { get { return "ObsoleteCallStatementInspection"; } }
        public string Description { get { return RubberduckUI.ObsoleteCall; } }
        public CodeInspectionType InspectionType { get { return CodeInspectionType.LanguageOpportunities; } }
        public CodeInspectionSeverity Severity { get; set; }

        public IEnumerable<CodeInspectionResultBase> GetInspectionResults(RubberduckParserState state)
        {
            return state.ObsoleteCallContexts.Select(context => 
                new ObsoleteCallStatementUsageInspectionResult(this,
                    new QualifiedContext<VBAParser.ExplicitCallStmtContext>(context.ModuleName, context.Context as VBAParser.ExplicitCallStmtContext)));
        }
    }
}