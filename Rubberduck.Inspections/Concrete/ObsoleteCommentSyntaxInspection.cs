using System.Collections.Generic;
using System.Linq;
using Antlr4.Runtime;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Results;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Inspections.Resources;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.Inspections.Concrete
{
    public sealed class ObsoleteCommentSyntaxInspection : InspectionBase, IParseTreeInspection
    {
        private IEnumerable<QualifiedContext> _parseTreeResults;
        private IEnumerable<QualifiedContext<VBAParser.RemCommentContext>> ParseTreeResults => _parseTreeResults.OfType<QualifiedContext<VBAParser.RemCommentContext>>();

        public ObsoleteCommentSyntaxInspection(RubberduckParserState state) : base(state, CodeInspectionSeverity.Suggestion) { }

        public override CodeInspectionType InspectionType => CodeInspectionType.LanguageOpportunities;

        public override IEnumerable<IInspectionResult> GetInspectionResults()
        {
            if (ParseTreeResults == null)
            {
                return Enumerable.Empty<IInspectionResult>();
            }
            return ParseTreeResults.Where(context => !IsIgnoringInspectionResultFor(context.ModuleName.Component, context.Context.Start.Line))
                .Select(context => new ObsoleteCommentSyntaxInspectionResult(this, new QualifiedContext<ParserRuleContext>(context.ModuleName, context.Context)));
        }

        public void SetResults(IEnumerable<QualifiedContext> results)
        {
            _parseTreeResults = results;
        }

        public class ObsoleteCommentSyntaxListener : VBAParserBaseListener
        {
            private readonly IList<VBAParser.RemCommentContext> _contexts = new List<VBAParser.RemCommentContext>();

            public IEnumerable<VBAParser.RemCommentContext> Contexts => _contexts;

            public override void ExitRemComment(VBAParser.RemCommentContext context)
            {
                _contexts.Add(context);
            }
        }
    }
}
