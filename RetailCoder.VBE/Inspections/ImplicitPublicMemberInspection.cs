using System.Collections.Generic;
using Antlr4.Runtime;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBA;
using Rubberduck.VBA.Nodes;
using Rubberduck.VBA.ParseTreeListeners;

namespace Rubberduck.Inspections
{
    public class ImplicitPublicMemberInspection : IInspection
    {
        public ImplicitPublicMemberInspection()
        {
            Severity = CodeInspectionSeverity.Suggestion;
        }

        public string Name { get { return InspectionNames.ImplicitPublicMember_; } }
        public CodeInspectionType InspectionType { get { return CodeInspectionType.MaintainabilityAndReadabilityIssues; } }
        public CodeInspectionSeverity Severity { get; set; }

        public IEnumerable<CodeInspectionResultBase> GetInspectionResults(VBProjectParseResult parseResult)
        {
            foreach (var module in parseResult.ComponentParseResults)
            {
                var procedures = module.ParseTree.GetContexts<ProcedureListener, ParserRuleContext>(new ProcedureListener(module.QualifiedName));
                foreach (var procedure in procedures)
                {
                    var context = (dynamic) procedure.Context;
                    if (((context.visibility() as VBAParser.VisibilityContext).GetAccessibility() == Accessibility.Implicit))
                    {
                        yield return new ImplicitPublicMemberInspectionResult(string.Format(Name,context.ambiguousIdentifier().GetText()), Severity, procedure);
                    }
                }
            }
        }
    }
}