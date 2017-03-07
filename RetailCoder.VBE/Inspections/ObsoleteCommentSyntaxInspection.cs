using System.Collections.Generic;
using System.Linq;
using Antlr4.Runtime;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Resources;
using Rubberduck.Inspections.Results;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.Inspections
{
    public sealed class ObsoleteCommentSyntaxInspection : InspectionBase, IParseTreeInspection<VBAParser.RemCommentContext>
    {
        private IEnumerable<QualifiedContext> _results;

        public ObsoleteCommentSyntaxInspection(RubberduckParserState state) : base(state, CodeInspectionSeverity.Suggestion) { }

        public override string Meta { get { return InspectionsUI.ObsoleteCommentSyntaxInspectionMeta; } }
        public override string Description { get { return InspectionsUI.ObsoleteCommentSyntaxInspectionName; } }
        public override CodeInspectionType InspectionType { get { return CodeInspectionType.LanguageOpportunities; } }

        public override IEnumerable<InspectionResultBase> GetInspectionResults()
        {
            if (ParseTreeResults == null)
            {
                return new InspectionResultBase[] { };
            }
            return ParseTreeResults.Where(context => !IsIgnoringInspectionResultFor(context.ModuleName.Component, context.Context.Start.Line))
                .Select(context => new ObsoleteCommentSyntaxInspectionResult(this, new QualifiedContext<ParserRuleContext>(context.ModuleName, context.Context)));
        }

        public void SetResults(IEnumerable<QualifiedContext> results)
        {
            _results = results;
        }

        public IEnumerable<QualifiedContext<VBAParser.RemCommentContext>> ParseTreeResults { get { return _results.OfType<QualifiedContext<VBAParser.RemCommentContext>>(); } }


        public class ObsoleteCommentSyntaxListener : VBAParserBaseListener
        {
            private readonly IList<VBAParser.RemCommentContext> _contexts = new List<VBAParser.RemCommentContext>();

            public IEnumerable<VBAParser.RemCommentContext> Contexts
            {
                get { return _contexts; }
            }

            public override void ExitRemComment(VBAParser.RemCommentContext context)
            {
                _contexts.Add(context);
            }
        }
    }
}
