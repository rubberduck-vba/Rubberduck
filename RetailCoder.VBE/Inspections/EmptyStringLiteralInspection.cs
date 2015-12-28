using System.Collections.Generic;
using System.Linq;
using Antlr4.Runtime;
using Rubberduck.Parsing;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.Inspections
{
    public class EmptyStringLiteralInspection : IInspection
    {
        public EmptyStringLiteralInspection()
        {
            Severity = CodeInspectionSeverity.Warning;
        }

        public string Name { get { return "EmptyStringLiteralInspection"; } }
        public string Description { get { return InspectionsUI.EmptyStringLiteralInspection; } }
        public CodeInspectionType InspectionType { get { return CodeInspectionType.LanguageOpportunities; } }
        public CodeInspectionSeverity Severity { get; set; }

        public IEnumerable<CodeInspectionResultBase> GetInspectionResults(RubberduckParserState state)
        {
            return
                state.EmptyStringLiterals.Select(
                    context =>
                        new EmptyStringLiteralInspectionResult(this,
                            new QualifiedContext<ParserRuleContext>(context.ModuleName, context.Context)));
        }
    }
}
