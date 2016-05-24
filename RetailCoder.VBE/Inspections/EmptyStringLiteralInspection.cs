using System.Collections.Generic;
using System.Linq;
using Antlr4.Runtime;
using Rubberduck.Parsing;
using Rubberduck.Parsing.VBA;
using Rubberduck.Parsing.Grammar;

namespace Rubberduck.Inspections
{
    public sealed class EmptyStringLiteralInspection : InspectionBase, IParseTreeInspection
    {
        public EmptyStringLiteralInspection(RubberduckParserState state)
            : base(state)
        {
        }

        public override string Meta { get { return InspectionsUI.EmptyStringLiteralInspectionMeta; } }
        public override string Description { get { return InspectionsUI.EmptyStringLiteralInspection; } }
        public override CodeInspectionType InspectionType { get { return CodeInspectionType.LanguageOpportunities; } }

        public ParseTreeResults ParseTreeResults { get; set; }

        public override IEnumerable<InspectionResultBase> GetInspectionResults()
        {   
            if (ParseTreeResults == null)
            {
                return new InspectionResultBase[] { };
            }
            return ParseTreeResults.EmptyStringLiterals.Select(
                    context => new EmptyStringLiteralInspectionResult(this,
                            new QualifiedContext<ParserRuleContext>(context.ModuleName, context.Context)));
        }

        public class EmptyStringLiteralListener : VBAParserBaseListener
        {
            private readonly IList<VBAParser.LiteralExpressionContext> _contexts = new List<VBAParser.LiteralExpressionContext>();
            public IEnumerable<VBAParser.LiteralExpressionContext> Contexts { get { return _contexts; } }

            public override void ExitLiteralExpression(VBAParser.LiteralExpressionContext context)
            {
                var literal = context.STRINGLITERAL();
                if (literal != null && literal.GetText() == "\"\"")
                {
                    _contexts.Add(context);
                }
            }
        }
    }
}
