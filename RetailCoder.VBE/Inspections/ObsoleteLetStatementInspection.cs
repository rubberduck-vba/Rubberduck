using System.Collections.Generic;
using System.Linq;
using Antlr4.Runtime;
using Rubberduck.Parsing;
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

        public IEnumerable<CodeInspectionResultBase> GetInspectionResults(VBProjectParseResult parseResult)
        {
            var issues = parseResult.Declarations.Items
                .Where(item => !item.IsBuiltIn)
                .SelectMany(item =>
                item.References.Where(reference => reference.HasExplicitLetStatement))
                .Select(issue => new ObsoleteLetStatementUsageInspectionResult(Description, Severity, new QualifiedContext<ParserRuleContext>(issue.QualifiedModuleName, issue.Context)));

            return issues;
        }
    }
}