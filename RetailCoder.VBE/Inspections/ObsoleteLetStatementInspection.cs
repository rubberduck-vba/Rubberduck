using System.Collections.Generic;
using System.Linq;
using Antlr4.Runtime;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Listeners;

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
            foreach (var result in parseResult.ComponentParseResults)
            {
                var module = result;
                var results = module.ParseTree.GetContexts<ObsoleteInstrutionsListener, ParserRuleContext>(new ObsoleteInstrutionsListener(module.QualifiedName))
                    .Select(context => context.Context)
                    .OfType<VBAParser.LetStmtContext>()
                    .Where(context => context.LET() != null && !string.IsNullOrEmpty(context.LET().GetText()))
                    .Select(context => new ObsoleteLetStatementUsageInspectionResult(Name, Severity, new QualifiedContext<VBAParser.LetStmtContext>(module.QualifiedName, context)));
                foreach (var inspectionResult in results)
                {
                    yield return inspectionResult;
                }
            }
        }
    }
}