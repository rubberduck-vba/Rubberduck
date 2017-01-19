using System.Collections.Generic;
using System.Linq;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Resources;
using Rubberduck.Inspections.Results;
using Rubberduck.Parsing;
using Rubberduck.Parsing.VBA;
using Rubberduck.Parsing.Grammar;

namespace Rubberduck.Inspections
{
    public sealed class EmptyStringLiteralInspection : InspectionBase, IParseTreeInspection<VBAParser.LiteralExpressionContext>
    {
        private IEnumerable<QualifiedContext> _parseTreeResults;

        public EmptyStringLiteralInspection(RubberduckParserState state)
            : base(state)
        {
        }

        public override string Meta { get { return InspectionsUI.EmptyStringLiteralInspectionMeta; } }
        public override string Description { get { return InspectionsUI.EmptyStringLiteralInspectionName; } }
        public override CodeInspectionType InspectionType { get { return CodeInspectionType.LanguageOpportunities; } }

        public IEnumerable<QualifiedContext<VBAParser.LiteralExpressionContext>> ParseTreeResults { get { return _parseTreeResults.OfType<QualifiedContext<VBAParser.LiteralExpressionContext>>(); } }

        public void SetResults(IEnumerable<QualifiedContext> results)
        {
            _parseTreeResults = results;
        }

        public override IEnumerable<InspectionResultBase> GetInspectionResults()
        {   
            if (ParseTreeResults == null)
            {
                return new InspectionResultBase[] { };
            }
            return ParseTreeResults
                .Where(result => !IsIgnoringInspectionResultFor(result.ModuleName.Component, result.Context.Start.Line))
                .Select(result => new EmptyStringLiteralInspectionResult(this, result));
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
