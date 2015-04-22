using System.Collections.Generic;
using System.Linq;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;

namespace Rubberduck.Inspections
{
    public class ObsoleteCallStatementInspection : IInspection
    {
        public ObsoleteCallStatementInspection()
        {
            Severity = CodeInspectionSeverity.Suggestion;
        }

        public string Name { get { return InspectionNames.ObsoleteCall; } }
        public CodeInspectionType InspectionType { get { return CodeInspectionType.LanguageOpportunities; } }
        public CodeInspectionSeverity Severity { get; set; }

        public IEnumerable<CodeInspectionResultBase> GetInspectionResults(VBProjectParseResult parseResult)
        {
            var issues = parseResult.Declarations.Items.SelectMany(declaration => 
                declaration.References.Where(reference => reference.HasExplicitCallStatement()))
                .Select(issue => new ObsoleteCallStatementUsageInspectionResult(Name, Severity,
                    new QualifiedContext<VBAParser.ExplicitCallStmtContext>(issue.QualifiedModuleName, issue.Context.Parent as VBAParser.ExplicitCallStmtContext)));

            return issues;
        }
    }
}