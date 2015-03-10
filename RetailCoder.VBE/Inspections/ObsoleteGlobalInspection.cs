using System.Collections.Generic;
using System.Linq;
using Antlr4.Runtime;
using Rubberduck.Parsing;
using Rubberduck.VBA;
using Rubberduck.VBA.Grammar;
using Rubberduck.VBA.Nodes;
using Rubberduck.VBA.ParseTreeListeners;

namespace Rubberduck.Inspections
{
    public class ObsoleteGlobalInspection : IInspection
    {
        public ObsoleteGlobalInspection()
        {
            Severity = CodeInspectionSeverity.Suggestion;
        }

        public string Name { get { return InspectionNames.ObsoleteGlobal; } }
        public CodeInspectionType InspectionType { get { return CodeInspectionType.LanguageOpportunities; } }
        public CodeInspectionSeverity Severity { get; set; }

        public IEnumerable<CodeInspectionResultBase> GetInspectionResults(VBProjectParseResult parseResult)
        {
            foreach (var result in parseResult.ComponentParseResults)
            {
                var statements = (result.ParseTree.GetContexts<DeclarationSectionListener, ParserRuleContext>(new DeclarationSectionListener(result.QualifiedName)))
                    .Select(context => context.Context).ToList();
                var module = result;
                foreach (var inspectionResult in
                    statements.OfType<VBAParser.VisibilityContext>()
                        .Where(context => context.GetText() == Tokens.Global)
                        .Select(context => new ObsoleteGlobalInspectionResult(Name, Severity, new QualifiedContext<ParserRuleContext>(module.QualifiedName, context.Parent as ParserRuleContext))))
                {
                    yield return inspectionResult;
                }
            }
        }
    }
}