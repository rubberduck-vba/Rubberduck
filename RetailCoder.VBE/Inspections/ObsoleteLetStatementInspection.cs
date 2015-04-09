using System.Collections.Generic;
using System.Linq;
using Antlr4.Runtime;
using Rubberduck.Parsing;

namespace Rubberduck.Inspections
{
    public class ObsoleteLetStatementInspection : IInspection
    {
        public ObsoleteLetStatementInspection()
        {
            Severity = CodeInspectionSeverity.Warning;
        }

        public string Name { get { return InspectionNames.ObsoleteLet; } }
        public CodeInspectionType InspectionType { get { return CodeInspectionType.LanguageOpportunities; } }
        public CodeInspectionSeverity Severity { get; set; }

        public IEnumerable<CodeInspectionResultBase> GetInspectionResults(VBProjectParseResult parseResult)
        {
            var issues = parseResult.Declarations.Items.SelectMany(item =>
                item.References.Where(reference => reference.HasExplicitLetStatement))
                .Select(issue => new ObsoleteLetStatementUsageInspectionResult(Name, Severity, new QualifiedContext<ParserRuleContext>(issue.QualifiedModuleName, issue.Context)));

            return issues;
        }
    }
}