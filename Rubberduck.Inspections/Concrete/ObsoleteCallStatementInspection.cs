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
    public sealed class ObsoleteCallStatementInspection : InspectionBase, IParseTreeInspection
    {
        private IEnumerable<QualifiedContext> _parseTreeResults;

        public ObsoleteCallStatementInspection(RubberduckParserState state)
            : base(state, CodeInspectionSeverity.Suggestion) { }

        public override CodeInspectionType InspectionType => CodeInspectionType.LanguageOpportunities;

        public void SetResults(IEnumerable<QualifiedContext> results)
        {
            _parseTreeResults = results;
        }

        public override IEnumerable<IInspectionResult> GetInspectionResults()
        {
            if (_parseTreeResults == null)
            {
                return new InspectionResultBase[] { };
            }

            var results = new List<ObsoleteCallStatementUsageInspectionResult>();

            foreach (var context in _parseTreeResults.Where(context => !IsIgnoringInspectionResultFor(context.ModuleName.Component, context.Context.Start.Line)))
            {
                var module = context.ModuleName.Component.CodeModule;
                var lines = module.GetLines(context.Context.Start.Line,
                    context.Context.Stop.Line - context.Context.Start.Line + 1);

                var stringStrippedLines = string.Join(string.Empty, lines).StripStringLiterals();

                int commentIndex;
                if (stringStrippedLines.HasComment(out commentIndex))
                {
                    stringStrippedLines = stringStrippedLines.Remove(commentIndex);
                }

                if (!stringStrippedLines.Contains(":"))
                {
                    results.Add(new ObsoleteCallStatementUsageInspectionResult(this,
                        new QualifiedContext<VBAParser.CallStmtContext>(context.ModuleName,
                            (VBAParser.CallStmtContext) context.Context)));
                }
            }

            return results;
        }

        public class ObsoleteCallStatementListener : VBAParserBaseListener
        {
            private readonly IList<VBAParser.CallStmtContext> _contexts = new List<VBAParser.CallStmtContext>();
            public IEnumerable<VBAParser.CallStmtContext> Contexts => _contexts;

            public override void ExitCallStmt(VBAParser.CallStmtContext context)
            {
                if (context.CALL() != null)
                {
                    _contexts.Add(context);
                }
            }
        }
    }
}
