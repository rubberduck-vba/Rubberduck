using System.Collections.Generic;
using System.Linq;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Results;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Inspections.Resources;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.Inspections.Concrete
{
    public sealed class EmptyStringLiteralInspection : InspectionBase, IParseTreeInspection
    {
        private IEnumerable<QualifiedContext> _parseTreeResults;

        public EmptyStringLiteralInspection(RubberduckParserState state)
            : base(state)
        {
        }

        public override CodeInspectionType InspectionType { get { return CodeInspectionType.LanguageOpportunities; } }

        public void SetResults(IEnumerable<QualifiedContext> results)
        {
            _parseTreeResults = results;
        }

        public override IEnumerable<IInspectionResult> GetInspectionResults()
        {   
            if (_parseTreeResults == null)
            {
                return Enumerable.Empty<IInspectionResult>();
            }
            return _parseTreeResults
                .Where(result => result.Context is VBAParser.LiteralExpressionContext 
                    && !IsIgnoringInspectionResultFor(result.ModuleName.Component, result.Context.Start.Line))
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
